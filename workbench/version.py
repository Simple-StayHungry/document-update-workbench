"""Single source of truth for the engine version boundary.

Everything else reads from here — the service payload, the exported report
headers, the launcher, and ``engine_manifest.json``. Nothing keeps its own copy
of the version string, so those places cannot drift apart again.
"""

VERSION = "1.20"
BUILD = "explain-workbench-" + VERSION
SCHEMA = 4
