export function* multiMatchs(re: RegExp, content: string) {
  let match: RegExpMatchArray | null;
  while ((match = re.exec(content)) !== null) {
    yield match;
  }
}
