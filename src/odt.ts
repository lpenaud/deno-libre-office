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

interface ListNodesOptions {
  start: number;
}

export interface FindNodesOptions {
  node: string;
  attributes?: NodeAttributes;
  start?: number;
}

export interface Node {
  node: string;
  attributes: NodeAttributes;
  index: number;
  endIndex: number;
  inner: string;
}

export type NodeType =
  | "text:p"
  | "text:span"
  | "table:table"
  | "table:table-row"
  | "table:table-cell"
  | "table:table-column"
  | "table:table-header-rows";

export interface CreateNodeOptions {
  node: NodeType;
  attributes: NodeAttributes;
  content: string;
  children: string[];
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
    find: { firstNode: FindNodesOptions; lastNode?: FindNodesOptions },
    inner: string,
  ) {
    const content = await this.readContent();
    const firstNode = findFirstNode(content, find.firstNode);
    if (firstNode === undefined) {
      console.warn("Node not found", find);
      return;
    }
    if (find.lastNode === undefined) {
      this.#unsafe(content, firstNode, inner);
      return;
    }
    const lastNode = findLastNode(content, {
      ...find.lastNode,
      start: firstNode.endIndex,
    });
    this.#unsafe(content, {
      index: firstNode.index,
      endIndex: lastNode === undefined ? firstNode.endIndex : lastNode.endIndex,
    }, inner);
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

  #unsafe(
    content: string,
    { index, endIndex }: { index: number; endIndex: number },
    inner: string,
  ) {
    this.#content = content.substring(0, index) + inner +
      content.substring(endIndex);
  }
}

export class NodeFactory implements CreateNodeOptions {
  #options: CreateNodeOptions;

  get node(): NodeType {
    return this.#options.node;
  }

  set node(node: NodeType) {
    this.#options.node = node;
  }

  get attributes(): NodeAttributes {
    return this.#options.attributes;
  }

  set attributes(attributes: NodeAttributes) {
    this.#options.attributes = attributes;
  }

  get content(): string {
    return this.#options.content;
  }

  set content(content: string) {
    this.#options.content = content;
  }

  get children(): string[] {
    return this.#options.children;
  }

  set children(children: string[]) {
    this.#options.children = children;
  }

  constructor(options: Partial<CreateNodeOptions> & { node: NodeType }) {
    this.#options = {
      attributes: {},
      children: [],
      content: "",
      ...options,
    };
  }

  newNode(options?: Partial<CreateNodeOptions> | string) {
    const { attributes, content, children, node }: CreateNodeOptions =
      typeof options === "string"
        ? { ...this.#options, content: options }
        : { ...this.#options, ...options, attributes: { ...this.#options.attributes, ...options?.attributes } };
    const attr = Object.entries(attributes)
      .map(([k, v]) => `${k}="${v}"`)
      .join(" ");
    return `<${node} ${attr}>${escapeLibreOffice(content)}${
      children.join("")
    }</${node}>`;
  }
}

export function* findNodes(
  xml: string,
  options: FindNodesOptions,
): Generator<Node> {
  const nodes = listNodes(xml, options);
  const startNode = findNodeStart(nodes, options);
  const endNode = findNodeEnd(nodes, options);
  for (;;) {
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

function findFirstNode(
  xml: string,
  options: FindNodesOptions,
): Node | undefined {
  for (const node of findNodes(xml, options)) {
    return node;
  }
}

function findLastNode(
  xml: string,
  options: FindNodesOptions,
): Node | undefined {
  let last: Node | undefined;
  for (const node of findNodes(xml, options)) {
    last = node;
  }
  return last;
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

function* listNodes(
  xml: string,
  options?: Partial<ListNodesOptions>,
): Generator<RawNode> {
  const { start }: ListNodesOptions = {
    start: 0,
    ...options,
  };
  const reNode = /<(\/)?(\w+:[\w-]+)\s*([!-~]+="[ -!#-;=?-~]+"\s*)*(\/)?>/gm;
  for (const match of multiMatchs(reNode, xml.substring(start))) {
    const [all, end, node, , prematureEnd] = match;
    const index = start + (match.index as number);
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
