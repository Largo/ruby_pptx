# ruby_pptx

Create, read and update PowerPoint (`.pptx`) files from Ruby.

A port of [python-pptx](https://github.com/scanny/python-pptx) — the same
battle-tested OOXML object model underneath, with a public API redesigned for
Ruby. See [PORTING.md](PORTING.md) for the architecture and the milestone plan.

> **Status: early.** The foundation layer (units, namespaces, part names) is in
> place and tested. Not yet usable for real work.

```ruby
require "ruby_pptx"

Pptx.inches(1).emu        #=> 914400
Pptx.cm(2.54).pt          #=> 72.0

require "pptx/core_ext"   # opt-in numeric sugar
1.inch == 72.points       #=> true
```

## Development

```bash
bundle install
bundle exec rspec
```

Specs compare the packages this gem writes against the ones python-pptx writes
for the same operation, part by part, on canonicalised XML. `python3 -c "import
pptx"` must work for those specs to run; they skip otherwise.

## License

MIT. python-pptx is MIT, © Steve Canny.
