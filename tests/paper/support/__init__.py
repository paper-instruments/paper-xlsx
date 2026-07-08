# Test-support helpers for the paper-xlsx contract harness (tests/paper).
#
# These are deliberately test-local for Phase 1. Once the package kernel
# (openpyxl.package, CONVENTIONS §7) lands in Phase 2, the semantic-XML and
# package-diff helpers here become thin wrappers over it; the assertions and
# fixture-handling helpers stay test-local forever.
