import * as path from "std/path/mod.ts";
import { multiMatchs } from "./utils.ts";

export type NodeAttributes = Record<string, string>;

interface RawNode {
  end: boolean;
  node: string;
  attributes: NodeAttributes;
  index: number;
  lastIndex: number;
}

export interface FindNodesOptions {
  node: string;
  attributes?: NodeAttributes;
}

export interface Node {
  node: string;
  attributes: NodeAttributes;
  index: number;
  endIndex: number;
  inner: string;
}

const LO_ENTITIES: Record<string, string> = {
  "&": "&amp",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;",
  "\t": "<text:tab />",
  "  ": "<text:tab />",
};
const LO_RE = new RegExp(Array.from(Object.keys(LO_ENTITIES)).join("|"), "g");

export class OpenDocumentText {
  #dir: string;

  #contentPath: string;

  #content: string | undefined;

  get dir() {
    return this.#dir;
  }

  static async fromZip(infile: string): Promise<OpenDocumentText> {
    const tempDir = await Deno.makeTempDir({
      prefix: "generate-odt",
    });
    const command = new Deno.Command("unzip", {
      args: ["-qd", tempDir, infile],
      stderr: "inherit",
      stdout: "inherit",
      stdin: "null",
    });
    const status = await command.spawn().status;
    if (status.success) {
      return new OpenDocumentText(tempDir);
    }
    throw new Error(`Unzip ends with ${status.code}`);
  }

  constructor(dir: string) {
    this.#dir = dir;
    this.#contentPath = path.join(dir, "content.xml");
  }

  async injectVar(name: string, value: string) {
    const content = await this.readContent();
    const node = findFirstNode(content, {
      node: "text:variable-set",
      attributes: {
        "text:name": name,
        "office:value-type": "string",
      },
    });
    if (node === undefined) {
      console.warn(`Cannot find variable "${name}"`);
      return;
    }
    this.#unsafe(
      content,
      node,
      `<text:variable-set text:name="${name}" office:value-type="string">${value}</text:variable-set>`,
    );
  }

  async injectVars(vars: Map<string, string>) {
    for (const [name, value] of vars) {
      // Need to be sync due to content manipulation
      await this.injectVar(name, value);
    }
  }

  async innerXml(
    find: { node: string; attributes?: Record<string, string> },
    inner: string,
  ) {
    const content = await this.readContent();
    const node = findFirstNode(content, find);
    if (node === undefined) {
      console.warn("Node not found", find);
      return;
    }
    this.#unsafe(content, node, inner);
  }

  async readContent() {
    if (this.#content === undefined) {
      this.#content = await Deno.readTextFile(this.#contentPath);
    }
    return this.#content;
  }

  async writeContent(filename: string = this.#contentPath) {
    return Deno.writeTextFile(filename, await this.readContent());
  }

  async save(filename: string) {
    const command = new Deno.Command("zip", {
      args: ["-q", "-9", "-r", path.resolve(filename), "."],
      cwd: this.#dir,
      stderr: "inherit",
      stdout: "inherit",
      stdin: "null",
    });
    await this.writeContent();
    const status = await command.spawn().status;
    if (!status.success) {
      throw new Error(`Zip ends with ${status.code}`);
    }
  }

  #unsafe(content: string, node: Node, inner: string) {
    this.#content = content.substring(0, node.index) + inner +
      content.substring(node.endIndex);
  }
}

export function* findNodes(
  xml: string,
  options: FindNodesOptions,
): Generator<Node> {
  const nodes = listNodes(xml);
  const startNode = findNodeStart(nodes, options);
  const endNode = findNodeEnd(nodes, options);
  for (; ;) {
    const { value: startValue } = startNode.next();
    if (startValue === undefined) {
      break;
    }
    const { value: endValue } = endNode.next();
    if (endValue === undefined) {
      break;
    }
    yield {
      node: startValue.node,
      attributes: startValue.attributes,
      index: startValue.index,
      endIndex: endValue.lastIndex,
      inner: xml.substring(startValue.index, endValue.lastIndex),
    };
  }
}

export function findFirstNode(
  xml: string,
  options: FindNodesOptions,
): Node | undefined {
  for (const node of findNodes(xml, options)) {
    return node;
  }
}

export function escapeLibreOffice(str: string) {
  return str.replace(LO_RE, (k) => LO_ENTITIES[k] as string);
}

function listAttributes(str: string): Record<string, string> {
  const re = /([!-~]+)="([ -!#-;=?-~]+)"/gm;
  const attributes: Record<string, string> = {};
  for (const [, key, value] of multiMatchs(re, str)) {
    attributes[key] = value;
  }
  return attributes;
}

function computeNodeStart(
  attributes?: NodeAttributes,
): (a: NodeAttributes) => boolean {
  if (attributes === undefined) {
    return () => true;
  }
  const entries = Object.entries(attributes);
  return (a) => entries.every(([k, v]) => a[k] === v);
}

function* findNodeStart(
  nodes: Iterable<RawNode>,
  { node, attributes }: FindNodesOptions,
) {
  const attrs = computeNodeStart(attributes);
  for (const match of nodes) {
    if (!match.end && match.node === node && attrs(match.attributes)) {
      yield match;
    }
  }
}

function* findNodeEnd(nodes: Iterable<RawNode>, { node }: FindNodesOptions) {
  for (const match of nodes) {
    if (match.end && match.node === node) {
      yield match;
    }
  }
}

function* listNodes(xml: string): Generator<RawNode> {
  const reNode = /<(\/)?(\w+:[\w-]+)\s*([!-~]+="[ -!#-;=?-~]+"\s*)*(\/)?>/gm;
  for (const match of multiMatchs(reNode, xml)) {
    const [all, end, node, , prematureEnd] = match;
    const index = match.index as number;
    yield {
      end: end !== undefined || prematureEnd !== undefined,
      node,
      attributes: listAttributes(all),
      index,
      lastIndex: index + all.length,
    };
  }
}

async function main(args: string[]): Promise<number> {
  if (args.length !== 1) {
    console.error("Usage %s XML", "odt");
    return 1;
  }
  const content = await Deno.readTextFile(args[0]);
  for (const node of findNodes(content, { node: "text:p" })) {
    console.log(content.substring(node.index, node.endIndex));
  }
  return 0;
}

if (import.meta.main) {
  main(Deno.args.slice())
    .then(Deno.exit);
}
