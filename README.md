# W5133 website

The website source lives on the `my-pages` branch. Email-management scripts live
in [W5133/email-management](https://github.com/W5133/email-management).

## Preview locally

From Terminal:

```sh
cd ~/Documents/git_projects/w5133.github.io
bundle install                 # first setup, or after Gemfile.lock changes
./scripts/preview
```

Open **http://127.0.0.1:4000/**. The server rebuilds when you save page edits, and
LiveReload refreshes the browser. Keep Terminal open; press **Control-C** to stop.
Restart the server after changing `_config.yml`.

For a build without starting a server:

```sh
./scripts/preview build
```

The generated site is written to `_site/`. Generated output and local caches are
ignored by Git. Neither command commits, pushes, or publishes anything.

## Requirements and troubleshooting

- Ruby and Bundler must be installed. This setup was verified with Ruby 3.2.2 and
  Bundler 2.6.9. `Gemfile.lock` records the tested dependency versions.
- The first setup downloads dependencies. Builds also download the configured
  Minima remote theme, so internet access is needed.
- If port 4000 is occupied, stop the existing preview or use
  `PREVIEW_PORT=4001 ./scripts/preview`, then open http://127.0.0.1:4001/.
- The older `2020-agenda.md` requests an unavailable `blog` layout and currently
  produces a build warning. It does not prevent previewing the other pages.

## Editing

- `index.md`: homepage and meeting overview.
- `2027-meeting.md`: upcoming meeting details.
- `past-meetings/` and `todo-boise.md`: archived meeting material.
- `assets/`: downloadable files, including the 2026 dinner scramble assignments.
- `_config.yml`: site title, theme, plugins, and navigation.

Before pushing, run the build and preview the pages you changed, checking their
links and layout. Publishing remains a separate Git push to `my-pages`.
