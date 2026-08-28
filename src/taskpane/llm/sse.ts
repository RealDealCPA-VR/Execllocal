/**
 * Minimal SSE helpers shared by the streaming transport.
 * NOTE: the SSE line prefix is assembled by concatenation to dodge a known
 * file-corruption hazard when the literal token appears in this repo.
 */
export const SSE_PREFIX = "dat" + "a:";
export const DONE_PAYLOAD = "[" + "DONE" + "]";

/**
 * Splits a (possibly partial) chunk of SSE text into complete payload lines.
 * Returns the remaining partial line so it can be prepended to the next chunk.
 */
export function extractSsePayloads(buffer: string): { payloads: string[]; rest: string } {
  const payloads: string[] = [];
  let rest = buffer;
  let idx = rest.indexOf("\n");
  while (idx >= 0) {
    const line = rest.slice(0, idx).trim();
    rest = rest.slice(idx + 1);
    if (line.startsWith(SSE_PREFIX)) {
      payloads.push(line.slice(SSE_PREFIX.length).trim());
    }
    idx = rest.indexOf("\n");
  }
  return { payloads, rest };
}
