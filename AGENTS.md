# Repository Guidelines

## Project Structure & Module Organization
`src/` contains the C application code, split by domain: `src/core` for startup and runtime services, `src/video` for streaming and detection, `src/storage` for recordings, `src/database` for SQLite access, and `src/web` for the embedded HTTP API. Public headers live in `include/` with the same module layout. Database migrations are in `db/migrations/`, runtime config samples are in `config/`, and longer design or operational notes are in `docs/`.

The browser UI lives in `web/`. Use `web/js/components/preact/` for UI components, `web/js/pages/` for page entry logic, and `web/css/` for styles. Playwright integration coverage is under `tests/integration/`; older Jest/Selenium web tests remain under `web/tests/`. Utility and release scripts are in `scripts/`.

## Build, Test, and Development Commands
Use the repo scripts instead of ad hoc commands when possible:

- `./scripts/build.sh` builds the C application in debug mode.
- `./scripts/build.sh --release` produces an optimized release build.
- `cmake -S . -B build/Release -DCMAKE_BUILD_TYPE=Release && cmake --build build/Release -j$(nproc)` is the manual build path.
- `./scripts/build_web_vite.sh` builds the Vite frontend into `web/dist/`.
- `npm test` runs Playwright integration tests from the repository root.
- `npm run test:api` or `npm run test:ui` runs tagged Playwright subsets.
- `cd web && npm test` runs frontend Jest tests.
- `ctest --test-dir build/Release --output-on-failure` runs compiled C tests after a CMake build.

## Coding Style & Naming Conventions
Match the surrounding code. C files use 4-space indentation, snake_case function names, and module-oriented filenames such as `api_handlers_system.c`. Frontend code uses 2-space indentation, ES modules, and PascalCase for Preact components like `StreamsPage.ts`. Keep new files within the existing domain directories instead of creating parallel structures. No formatter or linter config is committed here, so keep changes small and style-consistent.

## Testing Guidelines
Add C tests under `tests/` and register them in `tests/CMakeLists.txt`. Add browser or API integration coverage under `tests/integration/` using the existing `*.api.spec.ts` and `*.ui.spec.ts` naming. Reuse page objects from `tests/integration/pages/` for UI flows. For frontend-only logic, prefer Jest tests in `web/tests/`. Run the smallest relevant test set before opening a PR.

## Commit & Pull Request Guidelines
Recent history favors Conventional Commit prefixes: `feat:`, `fix:`, `refactor:`, `build:`, `perf:`. Continue that pattern and keep subjects imperative and specific. Pull requests should include a short problem statement, implementation summary, tests run, and screenshots or video for visible UI changes. Link related issues or docs when changing behavior, API responses, migrations, or deployment scripts.
