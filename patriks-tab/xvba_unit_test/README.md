# xVBA unit tests

`Test\Test.bas` contains source-level VBA checks for the FAS shortcut parser, shortcut catalog, and command registry.

XVBA watches this folder, but the official export flow includes `xvba_unit_test` in Production mode. To run the tests:

1. Use XVBA **Export VBA** and choose **Production**.
2. Open the XVBA macro list.
3. Run `Test.test`.

The test macros raise a VBA error on failure and show a pass message only after every check succeeds. They avoid workbook edits and live Excel smoke-test behavior.
