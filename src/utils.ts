export function* multiMatchs(re: RegExp, content: string) {
  if (!re.global) {
    throw new Error("RegExp haven't global flag");
  }
  let match: RegExpMatchArray | null;
  while ((match = re.exec(content)) !== null) {
    yield match;
  }
}

/**
 * Generate a array of string from start to end char code.
 * @param start Char code to start with.
 * @param end Char code to end with.
 */
export function* rangeChar(start: number, end: number) {
  for (;start <= end; start++) {
    yield String.fromCharCode(start);
  }
}
