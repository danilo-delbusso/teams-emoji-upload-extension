# Development

Set up the project as described in [Installation](./installation.md#from-source), then:

```console
npm run build         # build into dist/
npm test              # run unit tests with coverage
npm run format        # format with Prettier
npm run format:check  # check formatting, as CI does
npm run zip           # build and package chrome-extension.zip
```

After a build, click the reload icon on the extension's card in `chrome://extensions` to pick up the changes.

## Source layout

| File                                     | Purpose                                                     |
| ---------------------------------------- | ----------------------------------------------------------- |
| `popup.html`, `src/popup.ts`             | Side panel UI: upload list, settings, theme, view switching |
| `src/emojiBrowser.ts`                    | Browse and delete view, including the confirmation dialog   |
| `src/background.ts`                      | Service worker that runs uploads and reports progress       |
| `src/msTeams.ts`                         | Client for the Teams APIs                                   |
| `src/rateLimit.ts`                       | Handling Teams' 429 responses and pacing requests           |
| `src/draft.ts`                           | Saving the file list so it survives closing the panel       |
| `src/shortcuts.ts`, `src/customEmoji.ts` | Name defaults and validation, emoji list filtering          |

See [How it works](./how-it-works.md) for the Teams endpoints and the permissions the extension uses.
