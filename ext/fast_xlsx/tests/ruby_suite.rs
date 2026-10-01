// Runs the Ruby test suite against this crate's current source.
//
// The extension's behaviour is tested from Ruby (test/), but cargo-mutants only
// runs `cargo test`. This test rebuilds the extension from the (possibly
// mutated) source and runs the Ruby suite, so a mutant counts as caught when a
// Ruby test fails. The dev profile shares build artifacts with `cargo test`.
use std::path::Path;
use std::process::Command;

#[test]
fn ruby_test_suite_passes() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let status = Command::new("bundle")
        .args(["exec", "rake", "compile", "test"])
        .current_dir(&root)
        .env("RB_SYS_CARGO_PROFILE", "dev")
        // The suite needs a few CPU seconds; a mutant that loops forever is
        // stopped instead of running on after cargo-mutants gives up on it.
        .env("FAST_XLSX_TEST_CPU_SECONDS", "120")
        .status()
        .expect("could not run `bundle exec rake compile test`");
    assert!(status.success(), "the Ruby test suite failed");
}
