/**
 * PnP Parser Module
 * Handles parsing of Pick & Place data (Altium .txt format).
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

export class PnPParser {
  /**
   * Parse Altium Pick and Place .txt file
   * Supports space-delimited columns and quoted descriptions.
   */
  static parseAltium(text: string): PnPRow[] {
    const lines = text.split(/\r?\n/);
    const results: PnPRow[] = [];
    
    // 1. Find the header line (starts with Designator)
    const headerIdx = lines.findIndex(l => l.trim().startsWith('Designator') && l.includes('Description'));
    if (headerIdx === -1) return [];

    const header = lines[headerIdx];
    
    // Find column positions based on header labels
    const cols = [
      { name: 'designator', label: 'Designator' },
      { name: 'comment', label: 'Comment' },
      { name: 'layer', label: 'Layer' },
      { name: 'footprint', label: 'Footprint' },
      { name: 'x', label: 'Center-X' },
      { name: 'y', label: 'Center-Y' },
      { name: 'rotation', label: 'Rotation' },
      { name: 'description', label: 'Description' }
    ];

    const colPositions = cols.map(c => ({
      name: c.name,
      start: header.indexOf(c.label)
    })).filter(c => c.start !== -1).sort((a, b) => a.start - b.start);

    // 2. Parse data lines
    for (let i = headerIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line.trim() || line.startsWith('=')) continue;

      const row: Partial<PnPRow> = {};
      for (let j = 0; j < colPositions.length; j++) {
        const curr = colPositions[j];
        const next = colPositions[j + 1];
        
        // Extract substring
        let val = next 
          ? line.substring(curr.start, next.start) 
          : line.substring(curr.start);
        
        // Clean up: trim and remove quotes
        val = val.trim().replace(/^"|"$/g, '');
        row[curr.name as keyof PnPRow] = val;
      }

      if (row.designator) {
        results.push(row as PnPRow);
      }
    }

    return results;
  }
}
