/**
 * PnP Parser Module
 * Handles parsing of Pick & Place data (Altium .txt / .csv / .xlsx exports).
 */

export interface PnPRow {
  designator: string;
  description: string;
  comment?: string;
  layer?: string;
  footprint?: string;
  x?: string;
  y?: string;
  rotation?: string;
}

/** Split one delimited line, honoring double-quoted fields. */
function splitDelimited(line: string, delimiter: string): string[] {
  const out: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === delimiter) {
      out.push(field.trim());
      field = '';
    } else {
      field += ch;
    }
  }
  out.push(field.trim());
  return out;
}

/** Pick the delimiter that splits the header into the most fields. */
function detectDelimiter(header: string): string {
  const candidates = ['\t', ',', ';', '|'];
  let best = '';
  let bestCount = 1;
  for (const d of candidates) {
    const count = splitDelimited(header, d).length;
    if (count > bestCount) { best = d; bestCount = count; }
  }
  return best;
}

export class PnPParser {
  /**
   * Parse Altium Pick and Place files.
   * Supports tab/comma/semicolon-delimited files (quoted fields included) and
   * the legacy fixed-width .txt layout.
   */
  static parseAltium(text: string): PnPRow[] {
    const lines = text.split(/\r?\n/);

    // 1. Find the header line (starts with Designator)
    const headerIdx = lines.findIndex(l => {
      const t = l.trim();
      return t.startsWith('Designator') && t.includes('Description');
    });
    if (headerIdx === -1) return [];

    const header = lines[headerIdx];
    const delimiter = detectDelimiter(header);

    const cols: { name: keyof PnPRow; label: string }[] = [
      { name: 'designator', label: 'Designator' },
      { name: 'comment', label: 'Comment' },
      { name: 'layer', label: 'Layer' },
      { name: 'footprint', label: 'Footprint' },
      { name: 'x', label: 'Center-X' },
      { name: 'y', label: 'Center-Y' },
      { name: 'rotation', label: 'Rotation' },
      { name: 'description', label: 'Description' }
    ];

    const results: PnPRow[] = [];

    if (delimiter) {
      // 2a. Delimited layout — map header labels to field indexes
      const headerFields = splitDelimited(header, delimiter);
      const fieldFor = (label: string) =>
        headerFields.findIndex(f => f.trim().toLowerCase().startsWith(label.toLowerCase()));

      const indexes = cols.map(c => ({ name: c.name, index: fieldFor(c.label) }));
      const designatorIdx = indexes.find(i => i.name === 'designator')?.index ?? -1;

      for (let i = headerIdx + 1; i < lines.length; i++) {
        const line = lines[i];
        if (!line.trim() || line.startsWith('=')) continue;

        const fields = splitDelimited(line, delimiter);
        const row: Partial<PnPRow> = {};
        for (const { name, index } of indexes) {
          if (index !== -1) row[name] = fields[index] ?? '';
        }

        if (row.designator || (designatorIdx !== -1 && fields[designatorIdx])) {
          results.push(row as PnPRow);
        }
      }
      return results;
    }

    // 2b. Fixed-width layout — substring between header label offsets
    const colPositions = cols
      .map(c => ({ name: c.name, start: header.indexOf(c.label) }))
      .filter(c => c.start !== -1)
      .sort((a, b) => a.start - b.start);

    for (let i = headerIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line.trim() || line.startsWith('=')) continue;

      const row: Partial<PnPRow> = {};
      for (let j = 0; j < colPositions.length; j++) {
        const curr = colPositions[j];
        const next = colPositions[j + 1];
        const val = next
          ? line.substring(curr.start, next.start)
          : line.substring(curr.start);
        row[curr.name] = val.trim().replace(/^"|"$/g, '');
      }

      if (row.designator) results.push(row as PnPRow);
    }

    return results;
  }
}
