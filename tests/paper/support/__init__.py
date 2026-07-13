# Test-support helpers for the paper-xlsx contract harness (tests/paper).
#
# These are deliberately test-local. Once the package kernel
# (openpyxl.package) lands, the semantic-XML and
# package-diff helpers here become thin wrappers over it; the assertions and
# fixture-handling helpers stay test-local forever.
