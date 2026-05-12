# DummyExporter

Reference exporter plugin used to smoke-test officecli's plugin discovery and
to serve as a copy-pasteable starting point for third-party plugin authors. Not
part of the main solution; built independently when needed.

This fixture targets the synthetic `.test` extension so it doesn't conflict
with real exporters that users may have installed. Officecli has no built-in
view mode for `.test`, so end-to-end invocation through `view <file> <mode>`
isn't applicable — full coverage of the export path requires a plugin that
targets a real view mode (e.g. `view <file> pdf`).

## Build

```
dotnet publish examples/plugins/dummy-exporter -c Release -o examples/plugins/dummy-exporter/out
```

## Install

```
# Windows:
mkdir %USERPROFILE%\.officecli\plugins\exporter\test
copy examples\plugins\dummy-exporter\out\officecli-exporter-test.exe %USERPROFILE%\.officecli\plugins\exporter\test\plugin.exe

# Linux/macOS:
mkdir -p ~/.officecli/plugins/exporter/test
cp examples/plugins/dummy-exporter/out/officecli-exporter-test ~/.officecli/plugins/exporter/test/plugin
chmod +x ~/.officecli/plugins/exporter/test/plugin
```

## Verify discovery

```
officecli plugins list
# expect: officecli-exporter-test  0.1.0  exporter  .test  <path>

officecli plugins info officecli-exporter-test
# expect: full manifest including supports=["from:docx","from:xlsx","from:pptx"]
```

## What this fixture demonstrates

- `--info` manifest emission per docs/plugin-protocol.md §4
- The 4-path discovery resolution (it's installed at path #2, the user dir)
- Subprocess invocation surface per §5.2 (the `export <source> --out <target>`
  contract is implemented even though no main-side command currently dispatches
  to a `.test` target)
- Exit code conventions per §6.6 (0 success, 2 corrupt input, 64 invalid args)

For a complete end-to-end working example targeting a real format, see the
`view <file> pdf` path which dispatches to any installed exporter plugin
declaring `.pdf` in its manifest's `extensions` field.
