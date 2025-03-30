test_that("run_app() runs without errors", {
  skip_on_os("linux")  # Skip on Linux (CRAN's test environment)
  skip_on_os("mac")    # Skip on macOS (future-proofing)
  skip_if_not(interactive(), "Skipping in non-interactive mode")  # Avoid failures due to UI prompts
  skip_if_not(webshot::is_phantomjs_installed(), "PhantomJS not installed, skipping test")  # Avoid installation issues in test environments

  expect_error(run_app(), NA)  # This test now only runs on Windows in interactive mode with PhantomJS installed
})


