test_that("run_app() runs without errors", {
  skip_on_os("linux")  # Skip on Linux (CRAN's test environment)
  skip_on_os("mac")    # Skip on macOS (future-proofing)
  skip_if_not(interactive(), "Skipping in non-interactive mode")  # Avoid failures due to UI prompts

  expect_error(run_app(), NA)  # This test now only runs on Windows in interactive mode
})

