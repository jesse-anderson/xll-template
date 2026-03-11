# xll-template

A [cargo-generate](https://cargo-generate.github.io/cargo-generate/) template that scaffolds a fully working Excel XLL add-in project in Rust.

## What You Get

A generated project that:

- Compiles to a `.xll` file with `cargo build --release`
- Contains working example functions that appear in Excel immediately
- Includes export verification tests
- Optionally includes a GitHub Actions CI pipeline
- Requires zero manual configuration

## Prerequisites

- **Rust** (stable, MSVC toolchain on Windows) - [rustup.rs](https://rustup.rs)
- **cargo-generate** - install with `cargo install cargo-generate`
- **Windows** - XLLs are Windows DLLs, so building requires Windows or cross-compilation

## Quick Start

```sh
cargo generate gh:jesse-anderson/xll-template --name my-addin
```

You'll be prompted for:

| Prompt | Description | Default |
|--------|-------------|---------|
| **Display name** | Name shown in Excel's Add-In Manager | `My Add-In` |
| **Function prefix** | Prefix for Excel function names (e.g., `LR` produces `LR_ADD`) | *(empty)* |
| **Include CI** | Add a GitHub Actions workflow for building and releasing | `true` |

Or skip the prompts:

```sh
cargo generate gh:jesse-anderson/xll-template --name my-addin \
  --define addin_display_name="My Add-In" \
  --define function_prefix="LR" \
  --define include_ci=true
```

## Generated Project Structure

```
my-addin/
├── Cargo.toml          # Project manifest with xll-rs, xllgen, xll-utils deps
├── build.rs            # Build script - renames .dll to .xll
├── src/
│   └── lib.rs          # XLL source with example functions (ADD, MULTIPLY, VERSION)
├── tests/
│   └── verify_exports.rs  # Export verification tests
├── .github/workflows/  # (if CI enabled)
│   └── build.yml       # Build, test, and release workflow
├── .gitignore
└── README.md
```

## After Generating

### Build

```sh
cd my-addin
cargo build --release
```

The output is `target/release/my-addin.xll`.

### Test

```sh
cargo build --release
cargo test
```

### Load into Excel

1. Open Excel
2. Go to **File > Options > Add-ins**
3. At the bottom, select **Excel Add-ins** and click **Go...**
4. Click **Browse...** and select the `.xll` file
5. Click **OK**

Your functions (e.g., `=ADD(1, 2)` or `=LR_ADD(1, 2)` if you set a prefix) are now available in any cell.

## Built With

- [xll-rs](https://crates.io/crates/xll-rs) - XLL runtime
- [xllgen](https://crates.io/crates/xllgen) - Function registration macros
- [xll-utils](https://crates.io/crates/xll-utils) - Export verification

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE))
- MIT License ([LICENSE-MIT](LICENSE-MIT))

at your option.
