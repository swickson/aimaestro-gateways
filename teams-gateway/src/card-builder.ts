import { AdaptiveCard, TextBlock, FactSet, Fact } from '@microsoft/teams.cards';

export interface StatusSummaryFact {
  title: string;
  value: string;
}

export interface StatusSummary {
  title: string;
  status: 'success' | 'warning' | 'info' | 'error';
  description?: string;
  facts?: StatusSummaryFact[];
}

/**
 * Builds a typed Adaptive Card for status summary block.
 */
export function buildStatusSummaryCard(data: StatusSummary): AdaptiveCard {
  const elements = [];

  // Title
  elements.push(
    new TextBlock(data.title)
      .withSize('Large')
      .withWeight('Bolder')
      .withWrap(true)
  );

  // Status mapping to card element colors
  const colorMap = {
    success: 'Good',
    warning: 'Warning',
    error: 'Attention',
    info: 'Accent',
  } as const;

  elements.push(
    new TextBlock(`Status: ${data.status.toUpperCase()}`)
      .withWeight('Bolder')
      .withColor(colorMap[data.status] || 'Default')
      .withWrap(true)
  );

  // Description
  if (data.description) {
    elements.push(
      new TextBlock(data.description)
        .withWrap(true)
    );
  }

  // Facts
  if (data.facts && data.facts.length > 0) {
    const facts = data.facts.map((f) => new Fact(f.title, f.value));
    elements.push(new FactSet(...facts));
  }

  return new AdaptiveCard(...elements).withVersion('1.5');
}

/**
 * Unified card dispatcher at the SDK boundary.
 * Parses and converts structured content to an Adaptive Card object.
 * Returns a raw JS object representation of the card, or null if type is unsupported or parsing/validation fails.
 */
export function buildCard(type: string, messageText: string): Record<string, unknown> | null {
  if (type !== 'status_summary') {
    return null;
  }

  try {
    const data = JSON.parse(messageText) as StatusSummary;
    if (
      data &&
      typeof data === 'object' &&
      typeof data.title === 'string' &&
      typeof data.status === 'string' &&
      ['success', 'warning', 'info', 'error'].includes(data.status)
    ) {
      const card = buildStatusSummaryCard(data);
      // Serializes the builder classes into a plain JSON object matching standard schema
      return JSON.parse(JSON.stringify(card)) as Record<string, unknown>;
    }
  } catch {
    // Return null on parsing/validation failures to trigger fallback path
  }

  return null;
}
