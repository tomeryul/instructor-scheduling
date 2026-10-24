// SF Symbols-style glyphs drawn as inline SVG (24×24, currentColor).
// Filled variants are used in the tab bar, outline ones in bars and rows.

const P = {
  'calendar.fill': '<path d="M7 2.5a1 1 0 0 1 1 1V4h8v-.5a1 1 0 1 1 2 0V4h.5A2.5 2.5 0 0 1 21 6.5v12a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 18.5v-12A2.5 2.5 0 0 1 5.5 4H6v-.5a1 1 0 0 1 1-1zM5 9v9.5c0 .28.22.5.5.5h13a.5.5 0 0 0 .5-.5V9H5zm2.5 2.5h2v2h-2zm4.5 0h2v2h-2zm4.5 0h2v2h-2zm-9 4h2v2h-2zm4.5 0h2v2h-2z" fill="currentColor"/>',
  'person.2.fill': '<circle cx="9" cy="7.5" r="3.5" fill="currentColor"/><path d="M2 19c0-3.6 3.1-6 7-6s7 2.4 7 6c0 .6-.4 1-1 1H3c-.6 0-1-.4-1-1z" fill="currentColor"/><circle cx="17" cy="8.5" r="2.8" fill="currentColor" opacity=".75"/><path d="M17.5 13c2.7.2 4.5 2 4.5 4.6 0 .5-.3.9-.8.9h-3.6c.1-2.2-.6-4-2.1-5.3.6-.1 1.3-.2 2-.2z" fill="currentColor" opacity=".75"/>',
  'doc.fill': '<path d="M6.5 2h7.2c.66 0 1.3.26 1.77.73l3.8 3.8c.47.47.73 1.1.73 1.77V19.5A2.5 2.5 0 0 1 17.5 22h-11A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2zM13 3.8V7.5c0 .83.67 1.5 1.5 1.5h3.7L13 3.8zM8 13h8v1.6H8zm0 3.4h8V18H8z" fill="currentColor"/>',
  'folder.fill': '<path d="M3 6.5A2.5 2.5 0 0 1 5.5 4h3.9c.6 0 1.18.22 1.63.61L12.6 6H18.5A2.5 2.5 0 0 1 21 8.5v9a2.5 2.5 0 0 1-2.5 2.5h-13A2.5 2.5 0 0 1 3 17.5v-11z" fill="currentColor"/>',
  'ellipsis': '<circle cx="5.5" cy="12" r="1.9" fill="currentColor"/><circle cx="12" cy="12" r="1.9" fill="currentColor"/><circle cx="18.5" cy="12" r="1.9" fill="currentColor"/>',
  'plus': '<path d="M12 5v14M5 12h14" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>',
  'xmark': '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>',
  'checkmark': '<path d="M5 12.5l4.5 4.5L19 7.5" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>',
  'chevron.forward': '<path d="M9.5 5.5L16 12l-6.5 6.5" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>',
  'chevron.updown': '<path d="M8 9.5L12 5.5l4 4M8 14.5l4 4 4-4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>',
  'info.circle': '<circle cx="12" cy="12" r="9" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 11v6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="12" cy="7.6" r="1.2" fill="currentColor"/>',
  'magnifyingglass': '<circle cx="10.5" cy="10.5" r="6.5" fill="none" stroke="currentColor" stroke-width="2"/><path d="M15.5 15.5L20 20" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/>',
  'flask.fill': '<path d="M9 2.5h6a1 1 0 1 1 0 2h-.5v4.3l5.3 9.2A2.3 2.3 0 0 1 17.8 21.5H6.2A2.3 2.3 0 0 1 4.2 18l5.3-9.2V4.5H9a1 1 0 1 1 0-2zm2.5 2v4.8c0 .17-.05.35-.13.5L9.6 13h4.8l-1.77-3.2a1 1 0 0 1-.13-.5V4.5h-1z" fill="currentColor"/>',
  'mic.fill': '<rect x="8.5" y="2.5" width="7" height="12" rx="3.5" fill="currentColor"/><path d="M5.5 11a6.5 6.5 0 0 0 13 0M12 17.5V21M8.5 21h7" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'sparkles': '<path d="M10 3l1.7 4.8L16.5 9.5l-4.8 1.7L10 16l-1.7-4.8L3.5 9.5l4.8-1.7z" fill="currentColor"/><path d="M17.5 13l.9 2.6 2.6.9-2.6.9-.9 2.6-.9-2.6-2.6-.9 2.6-.9z" fill="currentColor"/>',
  'wand': '<path d="M4 20L14.5 9.5" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/><path d="M16.5 2.5l.9 2.4 2.4.9-2.4.9-.9 2.4-.9-2.4-2.4-.9 2.4-.9zM20 10l.6 1.4 1.4.6-1.4.6L20 14l-.6-1.4L18 12l1.4-.6z" fill="currentColor"/>',
  'exclamation.triangle.fill': '<path d="M10.3 3.5a2 2 0 0 1 3.4 0l7.9 13.6A2 2 0 0 1 19.9 20H4.1a2 2 0 0 1-1.7-2.9zM11 9v4.5a1 1 0 1 0 2 0V9a1 1 0 1 0-2 0zm1 6.6a1.2 1.2 0 1 0 0 2.4 1.2 1.2 0 0 0 0-2.4z" fill="currentColor"/>',
  'xmark.circle.fill': '<path d="M12 2.5a9.5 9.5 0 1 1 0 19 9.5 9.5 0 0 1 0-19zM9.2 7.8L7.8 9.2 10.6 12l-2.8 2.8 1.4 1.4L12 13.4l2.8 2.8 1.4-1.4-2.8-2.8 2.8-2.8-1.4-1.4L12 10.6z" fill="currentColor"/>',
  'plus.circle.fill': '<path d="M12 2.5a9.5 9.5 0 1 1 0 19 9.5 9.5 0 0 1 0-19zM11 7v4H7v2h4v4h2v-4h4v-2h-4V7z" fill="currentColor"/>',
  'arrow.clockwise': '<path d="M19 12a7 7 0 1 1-2.05-4.95" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19.5 4v4.5H15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>',
  'arrow.uturn': '<path d="M9 14L4.5 9.5 9 5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5 9.5h9.5a5 5 0 0 1 0 10H11" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>',
  'square.and.arrow.down': '<path d="M12 3.5v10M8 10l4 4 4-4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M8 7.5H6.5A2 2 0 0 0 4.5 9.5v9a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-9a2 2 0 0 0-2-2H16" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'square.and.arrow.up': '<path d="M12 14V3.5M8 7.5l4-4 4 4" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M8 9.5H6.5a2 2 0 0 0-2 2v7a2 2 0 0 0 2 2h11a2 2 0 0 0 2-2v-7a2 2 0 0 0-2-2H16" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'doc.on.doc': '<rect x="8" y="7" width="11.5" height="14" rx="2.5" fill="none" stroke="currentColor" stroke-width="1.9"/><path d="M5.5 16.5A2 2 0 0 1 4.5 15V5A2.5 2.5 0 0 1 7 2.5h7a2 2 0 0 1 1.7 1" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'trash': '<path d="M4.5 6.5h15M9.5 6.5V4.5a1 1 0 0 1 1-1h3a1 1 0 0 1 1 1v2M6.5 6.5l.9 12.6A2 2 0 0 0 9.4 21h5.2a2 2 0 0 0 2-1.9l.9-12.6M10 10.5v6.5M14 10.5v6.5" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>',
  'bell': '<path d="M6 16.5V11a6 6 0 0 1 12 0v5.5l1.5 2H4.5zM10 20.5a2 2 0 0 0 4 0" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linejoin="round" stroke-linecap="round"/>',
  'note': '<path d="M5.5 3.5h13a2 2 0 0 1 2 2v9l-6 6h-9a2 2 0 0 1-2-2v-13a2 2 0 0 1 2-2z" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linejoin="round"/><path d="M8 8.5h8M8 12h5" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'text.doc': '<path d="M5 4.5h14M5 9h14M5 13.5h9M5 18h11" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'eye.slash': '<path d="M3 12s3.3-6 9-6c1.5 0 2.8.4 4 1M21 12s-3.3 6-9 6c-1.5 0-2.8-.4-4-1" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><path d="M4 20L20 4" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><path d="M9.5 12a2.5 2.5 0 0 1 4-2" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'eye': '<path d="M2.5 12s3.5-6.5 9.5-6.5S21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linejoin="round"/><circle cx="12" cy="12" r="3" fill="none" stroke="currentColor" stroke-width="1.9"/>',
  'clock': '<circle cx="12" cy="12" r="8.8" fill="none" stroke="currentColor" stroke-width="1.9"/><path d="M12 7v5l3.3 2" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>',
  'person.badge.plus': '<circle cx="10" cy="8" r="3.6" fill="none" stroke="currentColor" stroke-width="1.9"/><path d="M3.5 19.5c.5-3.5 3.2-5.5 6.5-5.5 1.3 0 2.5.3 3.5.9" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><path d="M18 13.5v6M15 16.5h6" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/>',
  'tray.and.arrow.down': '<path d="M12 3v9.5M8.5 9l3.5 3.5L15.5 9" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/><path d="M3.5 14h5a3.5 3.5 0 0 0 7 0h5v4.5a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2z" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linejoin="round"/>',
  'tablecells': '<rect x="3.5" y="4.5" width="17" height="15" rx="2.5" fill="none" stroke="currentColor" stroke-width="1.9"/><path d="M3.5 9.5h17M3.5 14.5h17M9.5 9.5v10" stroke="currentColor" stroke-width="1.6"/>',
  'pdf': '<path d="M6.5 2.5h7l5 5v12a2 2 0 0 1-2 2h-10a2 2 0 0 1-2-2v-15a2 2 0 0 1 2-2z" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linejoin="round"/><path d="M13.5 2.5v5h5M8 13h8M8 16.5h5" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>',
  'slider': '<path d="M4 7h9M17 7h3M4 17h3M11 17h9" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><circle cx="15" cy="7" r="2.2" fill="none" stroke="currentColor" stroke-width="1.9"/><circle cx="9" cy="17" r="2.2" fill="none" stroke="currentColor" stroke-width="1.9"/>',
};

export function icon(name, cls = '') {
  const body = P[name];
  if (!body) return '';
  return `<svg class="sym${cls ? ' ' + cls : ''}" viewBox="0 0 24 24" aria-hidden="true" focusable="false">${body}</svg>`;
}
