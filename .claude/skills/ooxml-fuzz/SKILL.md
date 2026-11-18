---
name: ooxml-fuzz
description: Fuzz ruby_pptx's writer with seeded random decks and check every saved .pptx with the Open XML SDK's OpenXmlValidator (.NET), then minimize, triage and fix what it finds. Use whenever the user wants to fuzz ruby_pptx, check that its output is valid OOXML / schema-valid / opens in PowerPoint without the repair dialog, run OpenXmlValidator or the Open XML SDK on generated decks, hunt for writer bugs before a release, or after changing anything that writes XML (oxml, chart XML, fills, text, tables) — even if they only say "check the output is correct".
---

# OOXML fuzzing for ruby_pptx

Random decks, built through the public API from a seeded RNG, saved, and
handed to `DocumentFormat.OpenXml`'s `OpenXmlValidator` — the same schema and
semantic checks as Microsoft's Open XML SDK Productivity Tool. The spec suite
proves parity with python-pptx; this proves the output is valid OOXML, and
finds the call sequences nobody wrote a spec for.

Everything lives in this directory:

| Path | What |
|---|---|
| `scripts/fuzz.rb` | generator, executor, runner, replay, minimizer (one file) |
| `validator/` | the .NET console app around `OpenXmlValidator` |
| `bin/` | its build output (gitignored, built on first use) |
| `known-issues.txt` | findings accepted on purpose, one regexp each, with the reason |

All commands run from the repo root with `bundle exec`. Output goes to
`tmp/ooxml-fuzz/` (gitignored) unless `--out` says otherwise.

## Prerequisites

A .NET SDK 8 or newer. The script finds `~/.dotnet/dotnet`, `/usr/bin/dotnet`,
`/usr/share/dotnet/dotnet` or `$DOTNET`. If none exists, install one first
(`https://dot.net/v1/dotnet-install.sh --channel 10.0`). The first run builds
the validator, which downloads `DocumentFormat.OpenXml` from NuGet; it
rebuilds itself whenever a file in `validator/` changes.

## Workflow

### 1. Run

```bash
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb run --count 200
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb run --count 200 --edge
```

- `--seed N` fixes the master seed; it is printed on every run, so rerun
  with it to check a fix against the same decks.
- `--ops N` caps the operations per deck (default 12).
- `--edge` adds hostile values: negative and overflowing EMUs, NaN and
  Infinity in chart data, control characters and U+FFFE in text, bad hex
  colours, out-of-range enums. With those, a rejection (`ArgumentError` and
  friends) is the right outcome; invalid XML is not.
- `RUBY_PPTX_XML_BACKEND=rexml` runs the same thing on the REXML backend.
- `OOXML_VERSION=Office2007` (or `Office2010` ... `Microsoft365`, the
  default) picks the file format version the validator checks against.

200 decks take about 20 seconds. The run prints distinct findings, grouped
by signature with a count and one example case; exceptions; and the path to
`report.json`, which also counts rejected ops. Exit status is non-zero when
there is a new finding or a library crash.

### 2. Read the report

Each deck is `case-<seed>.pptx` beside `case-<seed>.json`. The JSON holds the
ops that built it and a per-op log: `ok` (with notes on any sub-step that was
rejected or does not apply), `rejected`, or `crash`.

There are three kinds of result, and they mean different things:

1. **Validator finding**: the library wrote invalid OOXML. Always a bug, unless
   the input was itself invalid. Then the bug is that the library accepted it
   instead of raising `ArgumentError`.
   - `Sch_*`: schema violations (wrong child order, missing required
     attribute, out-of-range value, empty element that needs content).
   - `ExceptionError ... invalid character`: the part is not even well-formed
     XML (raw control characters). PowerPoint refuses the whole file.
   - `OPEN_ERROR`: the package could not be opened at all (broken
     relationship, content type, or zip).
2. **Crash, `library`**: an exception that is not a deliberate rejection,
   raised from inside `lib/` (a `NoMethodError` on nil, a Nokogiri
   `SyntaxError` while the library re-parses its own XML, an `ArgumentError`
   for the wrong number of arguments). A bug: the
   library should either succeed or raise `ArgumentError`/`Pptx::Error`.
3. **Crash, `harness?`**, or a sub-step logged as `n/a`: the fuzzer called
   something that does not exist. That is a fuzzer bug. Fix `fuzz.rb` and do
   not report it. `n/a` on a method that does not apply to that object, such
   as `marker` on a bar series, is expected.

A finding can come from an op that was **rejected**: the library raised, but
had already changed the XML (for example `theme_color = :MIXED` leaves a
`<a:schemeClr>` without `val`). That is a real bug too. A rejected call must
leave the document as it was, so validate before mutating.

### 3. Minimize

```bash
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb minimize tmp/ooxml-fuzz/case-<seed>.json \
  --signature '<signature as printed by run>'
```

This removes whole ops first (ddmin), then any hash key, array element or value
at any depth, until nothing more can go. It keeps one validator process running,
so a few hundred builds take seconds. It writes `case-<seed>-min.json`.
Without `--signature`, it targets the case's first finding. Quote the
signature exactly as `run` printed it.

The minimized case reproduces the same *signature*, which is not always the
same root cause: dropping a value can expose a neighbouring bug that fails
the same way. Read the minimized ops and check that they describe the bug
you are after.

### 4. Replay and inspect

```bash
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb replay tmp/ooxml-fuzz/case-<seed>-min.json
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb replay tmp/ooxml-fuzz/case-<seed>.json --keep 0,3
unzip -p tmp/ooxml-fuzz/case-<seed>-min-replay.pptx ppt/slides/slide1.xml | xmllint --format - | less
bundle exec ruby .claude/skills/ooxml-fuzz/scripts/fuzz.rb validate some.pptx   # any deck, from anywhere
```

The validator's `path` (an XPath such as `/p:sld[1]/.../a:pattFill[1]/a:bgClr[1]`)
points at the bad element. Compare it with ISO/IEC 29500-1 and with what
python-pptx writes for the same call. `tools/*_oracle.py` shows how the specs
drive python-pptx.

Then write the bug as a plain Ruby snippet against the public API, a few
lines with no fuzzer involved. That snippet is the repro for the issue and
the basis of the regression spec.

### 5. Fix

- Fix it in `lib/`. Keep python-pptx parity, which the differential specs
  enforce. If python-pptx writes the same invalid XML, say so: the fix is then
  a deliberate divergence and goes in PORTING.md's list of differences.
- Add a regression spec to `spec/ruby_pptx/ooxml_validity_spec.rb`, which
  collects the fuzzer's findings by area. Assert the XML now written, or the
  error now raised and that the XML is unchanged. Where the fix departs from
  python-pptx, also pin python-pptx's output in that file's last block and
  add a row to PORTING.md's "Upstream bugs this port does not reproduce". Where a
  slide-master-style schema check fits, `spec/support/schema.rb`
  (`schema_violations(path)`, needs `OOXML_SCHEMAS`) validates against the
  XSDs.
- Run `bundle exec rspec` and `bundle exec rubocop`.
- Rerun the fuzzer with the same master seed and `--edge` flag. The signature
  should be gone and nothing new should appear. Then run a few fresh seeds.

### 6. Accept on purpose, rarely

When a finding is deliberate, add a regexp to
`known-issues.txt` with a comment that says why it is accepted and what would
remove it. `run` then lists it under "known issue signatures suppressed"
instead of failing. Never suppress a finding just to get a green run.

## Extending the generator

Coverage is only as wide as the generator. When a feature lands, or a run
shows that part of the API never gets exercised, add it:

1. `Generator#gen_<name>` returns a plain JSON-able hash: strings, numbers,
   booleans, nil, arrays, hashes. Symbols become strings on the JSON
   round trip, so store enum names as strings and call `.to_sym` in the
   executor. Use `@edge` for the hostile variants.
2. `Executor#op_<name>(op)` applies it. Wrap each optional step in
   `try("label") { ... }`, so a rejected value is noted without dropping the
   rest of the op.
3. Give it a weight in `Generator::OP_WEIGHTS`.
4. Run 200 decks and check `report.json` for `harness?` crashes and `n/a`
   notes on your new op.

Existing case files keep replaying after the generator changes, because the
op list is what is stored. A *seed* only reproduces the same deck with the
same `fuzz.rb`.

## Limits

- `OpenXmlValidator` is not PowerPoint. A file it passes can still trigger
  the repair dialog (dangling relationship ids inside extLst, duplicate
  shape ids, inconsistent chart caches). A file it fails sometimes opens
  fine, because PowerPoint reads more than the schema allows. When it
  matters, open the minimized deck in PowerPoint.
- The validator checks the XML parts and package structure. It does not
  open the embedded chart workbooks or OLE payloads, and it does not check
  images.
- The fuzzer only uses fixtures under `spec/fixtures/`, plus DejaVuSans for
  `fit_text` (that op is rejected when the font is missing).
