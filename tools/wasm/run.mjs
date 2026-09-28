// Runs a Ruby script under ruby.wasm (Ruby 4.0) with the WASI layer the
// browser build uses -- @bjorn3/browser_wasi_shim and its in-memory
// filesystem -- so this is what a web page gets, minus the DOM. Node's own
// WASI is not used: it aborts on nested requires from mounted directories.
//
//   node run.mjs SCRIPT OUT_DIR NAME=LIB_DIR...
//
// Each LIB_DIR (a gem's lib/, template files and all) is copied into the
// in-memory filesystem at /gems/NAME and put on the load path. Nokogiri is
// never among them: it is native and cannot run in ruby.wasm. The script
// runs with /work holding the fixtures below; anything it writes to /work
// is copied to OUT_DIR afterwards.
import { mkdir, readFile, readdir, stat, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { ConsoleStdout, Directory, File, OpenFile, PreopenDirectory, WASI } from "@bjorn3/browser_wasi_shim";
import { RubyVM } from "@ruby/wasm-wasi/dist/vm";

const here = dirname(fileURLToPath(import.meta.url));
const [script, outDir, ...libs] = process.argv.slice(2);

async function tree(dir) {
  const entries = new Map();
  for (const name of await readdir(dir)) {
    const full = join(dir, name);
    entries.set(name, (await stat(full)).isDirectory() ? new Directory(await tree(full)) : new File(await readFile(full)));
  }
  return entries;
}

const gems = new Map();
for (const lib of libs) {
  const [name, path] = lib.split("=");
  gems.set(name, new Directory(await tree(path)));
}
const work = new Map([["logo.png", new File(await readFile(join(here, "../../spec/fixtures/images/png-96dpi.png")))]]);
const root = new Map([["gems", new Directory(gems)], ["work", new Directory(work)]]);

const fds = [
  new OpenFile(new File([])),
  ConsoleStdout.lineBuffered((line) => process.stdout.write(`${line}\n`)),
  ConsoleStdout.lineBuffered((line) => process.stderr.write(`${line}\n`)),
  new PreopenDirectory("/", root),
];
const wasi = new WASI([], [], fds, { debug: false });
const wasm = await readFile(join(here, "node_modules/@ruby/4.0-wasm-wasi/dist/ruby+stdlib.wasm"));
const { vm } = await RubyVM.instantiateModule({ module: await WebAssembly.compile(wasm), wasip1: wasi });

try {
  const loadPath = [...gems.keys()].map((name) => `"/gems/${name}"`).join(", ");
  vm.eval(`$LOAD_PATH.unshift(${loadPath})`);
  vm.eval(await readFile(script, "utf8"));
} catch (error) {
  process.stderr.write(`${error?.message ?? error}\n`);
  process.exitCode = 1;
}

await mkdir(outDir, { recursive: true });
for (const [name, entry] of work) {
  if (entry instanceof File) await writeFile(join(outDir, name), entry.data);
}
