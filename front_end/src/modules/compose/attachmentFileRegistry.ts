// filepath: my-app/src/modules/compose/attachmentFileRegistry.ts
// Module-scoped map of attachment ID -> raw File object. Kept in a
// dedicated file so the React Refresh rule (`only-export-components`)
// does not flag the uploader component file for also exporting a
// non-component value.

export const fileRegistry = new Map<string, File>();
