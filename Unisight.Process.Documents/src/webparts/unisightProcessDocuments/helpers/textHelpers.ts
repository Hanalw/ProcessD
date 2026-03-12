export const formatRichText = (input: string) => {
	// Return plain text from a rich HTML string.
	if (!input) return '';

	const normalize = (s: string) =>
		s.replace(/\u00A0/g, ' ') // convert non-breaking spaces
      .replace(/\s+/g, ' ')   // collapse whitespace
      .trim();

	try {
		const parser = typeof DOMParser !== 'undefined' ? new DOMParser() : null;
		if (parser) {
			const doc = parser.parseFromString(input, 'text/html');
			// Remove non-content elements
			doc.body.querySelectorAll('script, style, noscript, iframe').forEach(el => el.remove());
			// innerText preserves line breaks better than textContent
			const text = (doc.body as HTMLElement).innerText || doc.body.textContent || '';
			return normalize(text);
		}
	} catch {
		// ignore and use fallback
	}

	// Fallback: strip tags via regex
	const noTags = input.replace(/<[^>]+>/g, ' ');
	return normalize(noTags);
}

export const trimText = (text: string, maxLength: number) => {
	if (!text || maxLength < 1) return '';
	const trimmed = text.length > maxLength ? text.slice(0, maxLength - 1) + '…' : text;
	return trimmed;
}