import * as path from "std/path/mod.ts";
import { multiMatchs, rangeChar } from "./utils.ts";

export type NodeAttributes = Record<string, string | undefined>;

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

export type NodeTypeTable =
  | "table:table"
  | "table:table-row"
  | "table:table-cell"
  | "table:table-column"
  | "table:table-header-rows"
  | "table:covered-table-cell";

export type NodeTypeText =
  | "text:p"
  | "text:span";

export type NodeType = NodeTypeTable | NodeTypeText;

export interface CreateNodeOptions {
  node: NodeType;
  attributes: NodeAttributes;
  content: string;
  children: string[];
}

export interface OpenDocumentTextJson {
  odt: string;
  content: string;
}

const LO_ENTITIES: Record<string, string> = {
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&apos;",
  "\t": "<text:tab />",
  "  ": "<text:tab />",
};
const LO_RE = new RegExp(Object.keys(LO_ENTITIES).join("|"), "g");
const TEMP_DIR = await Deno.makeTempDir({
  prefix: "deno-lo",
});

async function readAll(readable: ReadableStream<Uint8Array>) {
  let content = "";
  for await (const buf of readable.pipeThrough(new TextDecoderStream())) {
    content += buf;
  }
  return content;
}

export function escapeLibreOffice(str: string) {
  return str.replace(LO_RE, (k) => LO_ENTITIES[k] as string);
}

export function removeTempDir() {
  return Deno.remove(TEMP_DIR, {
    recursive: true,
  });
}

/**
 * @param tableName Table name
 * @param columnDef Column definition
 * @returns The table columns list.
 */
export function createColumns(tableName: string, columnDef: string) {
  const factory = new NodeFactory({
    node: "table:table-column",
  });
  return calcColumns(columnDef).map((c) =>
    factory.newNode({
      attributes: {
        "table:style-name": `${tableName}.${c}`,
      },
    })
  );
}

export class OpenDocumentText {
  #odt: string;

  #content: string;

  static async fromZip(infile: string): Promise<OpenDocumentText> {
    const command = new Deno.Command("unzip", {
      args: ["-qp", infile, "content.xml"],
      stderr: "inherit",
      stdout: "piped",
      stdin: "null",
    });
    const process = command.spawn();
    const [status, content] = await Promise.all([
      process.status,
      readAll(process.stdout),
    ]);
    if (status.success) {
      return new OpenDocumentText({ odt: infile, content });
    }
    throw new Error(`Unzip ends with ${status.code}`);
  }

  constructor({ odt, content }: OpenDocumentTextJson) {
    this.#content = content;
    this.#odt = odt;
  }

  injectVar(name: string, value: string) {
    const attributes: NodeAttributes = {
      "text:name": name,
    };
    for (
      const node of findNodes(this.#content, {
        node: "text:variable-set",
        attributes,
      })
    ) {
      this.#unsafe(
        node,
        `<text:variable-set text:name="${name}" office:value-type="string">${value}</text:variable-set>`,
      );
    }
    for (
      const node of findNodes(this.#content, {
        node: "text:variable-get",
        attributes,
      })
    ) {
      this.#unsafe(
        node,
        `<text:variable-get text:name="${name}">${value}</text:variable-get>`,
      );
    }
  }

  injectVars(vars: Map<string, string>) {
    for (const [name, value] of vars) {
      this.injectVar(name, value);
    }
  }

  innerXml(
    find: { firstNode: FindNodesOptions; lastNode?: FindNodesOptions },
    inner: string,
  ) {
    const firstNode = findFirstNode(this.#content, find.firstNode);
    if (firstNode === undefined) {
      console.warn("Node not found", find);
      return;
    }
    if (find.lastNode === undefined) {
      this.#unsafe(firstNode, inner);
      return;
    }
    const lastNode = findLastNode(this.#content, {
      ...find.lastNode,
      start: firstNode.endIndex,
    });
    this.#unsafe({
      index: firstNode.index,
      endIndex: lastNode === undefined ? firstNode.endIndex : lastNode.endIndex,
    }, inner);
  }

  async save(filename: string) {
    filename = path.resolve(filename);
    const tempDir = await makeTempDir(path.basename(filename));
    const command = new Deno.Command("zip", {
      args: ["-q", "-9", path.resolve(filename), "content.xml"],
      cwd: tempDir,
      stderr: "inherit",
      stdout: "inherit",
      stdin: "null",
    });
    await Promise.all([
      // Write destination
      Deno.copyFile(this.#odt, filename),
      // Write content.xml in tempdir
      Deno.writeTextFile(path.join(tempDir, "content.xml"), this.#content),
    ]);
    const status = await command.output();
    if (!status.success) {
      throw new Error(`Zip ends with ${status.code}`);
    }
  }

  /**
   * Clone the current instance.
   * @returns Clone
   */
  clone(): OpenDocumentText {
    return new OpenDocumentText(this.toJson());
  }

  toJson(): OpenDocumentTextJson {
    return {
      odt: this.#odt,
      content: this.#content,
    };
  }

  #unsafe(
    { index, endIndex }: { index: number; endIndex: number },
    inner: string,
  ) {
    this.#content = this.#content.substring(0, index) + inner +
      this.#content.substring(endIndex);
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
      typeof options === "string" ? { ...this.#options, content: options } : {
        ...this.#options,
        ...options,
        attributes: { ...this.#options.attributes, ...options?.attributes },
      };
    const attr = Object.entries(attributes)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `${k}="${v}"`)
      .join(" ");
    let result = `<${node}`;
    if (attr.length > 0) {
      result += ` ${attr}`;
    }
    if (content.length === 0 && children.length === 0) {
      return `${result}/>`;
    }
    return `${result}>${escapeLibreOffice(content)}${
      children.join("")
    }</${node}>`;
  }
}

/**
 * Create the columns list
 * @example
 * calcColumns("A-C")
 * // ["A", "B", "C"]
 * calcColumns("E")
 * // ["E"]
 * calcColumns("A-BB-E")
 * // ["A", "B", "B", "C", "D", "E"]
 * @param columnDef Column definition
 * @returns List of columns names.
 */
function calcColumns(columnDef: string) {
  if (columnDef.length === 1) {
    return [columnDef];
  }
  const stringFactory: (char: number) => string = (char) =>
    String.fromCharCode(char);
  const result: string[] = [];
  let hyphen = false;
  let group: number[] = [];
  for (let i = 0; i < columnDef.length; i++) {
    const char = columnDef.charCodeAt(i);
    // 'A' >= char <= 'Z'
    if (char >= 65 && char <= 90) {
      if (group.push(char) >= 2 && hyphen) {
        const end = group.pop() as number;
        const start = group.pop() as number;
        result.push(...group.map(stringFactory), ...rangeChar(start, end));
        group = [];
      }
      hyphen = false;
      continue;
    }
    // char === '-'
    if (char === 45) {
      hyphen = true;
    }
  }
  if (group.length > 0) {
    result.push(...group.map(stringFactory));
  }
  return result;
}

function* findNodes(
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

function listAttributes(str: string): NodeAttributes {
  const re = /([!-~]+)="([ -!#-;=?-~]+)"/gm;
  const attributes: NodeAttributes = {};
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

function makeTempDir(prefix?: string) {
  return Deno.makeTempDir({
    dir: TEMP_DIR,
    prefix,
  });
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
