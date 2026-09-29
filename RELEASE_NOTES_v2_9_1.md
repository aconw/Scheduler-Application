# v2.9.1 Release Notes

## Upload validation compatibility fix
The Streamlit upload validator now calls `read_sheet()` with explicit `required_headers=` and `optional_headers=` keyword arguments. This prevents an older compatible engine signature from interpreting the expected header list as a worksheet name.

No scheduling business rules were changed in v2.9.1.
