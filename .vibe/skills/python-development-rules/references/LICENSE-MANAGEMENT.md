# License Management

## Discovery

Before writing any code, detect the project's license:

1. Check `pyproject.toml` field `tool.poetry.license` or `project.license`
2. Check for a `LICENSE` or `LICENSE.txt` file at the project root
3. Check `AGENTS.md` or `README.md` for license mentions
4. If no license is found, **ask the user** before proceeding

## License Rules

### GPL v3 (GNU General Public License v3)

**File header**: Every Python file must start with:

```python
# Copyright (C) <year> <author>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.
```

**Constraints**:
- All derivative works must also be GPL v3
- Source code must be made available with any distribution
- Changes must be documented

### MIT License

**File header**: Optional but recommended:

```python
# Copyright (C) <year> <author>
# SPDX-License-Identifier: MIT
```

**Constraints**:
- Include copyright notice and license in all copies
- No restrictions on use, modification, or distribution

### Apache 2.0

**File header**: Required:

```python
# Copyright <year> <author>
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
```

**Constraints**:
- Include NOTICE file if one exists
- State changes made to original code
- Patent grant included

### BSD (2-Clause / 3-Clause)

**File header**: Optional but recommended:

```python
# Copyright (C) <year> <author>
# SPDX-License-Identifier: BSD-2-Clause
```

**Constraints**:
- Include copyright notice in source and binary distributions
- 3-Clause adds: no use of author name for endorsement without permission

### Proprietary / No License

**File header**: Required:

```python
# Copyright (C) <year> <author>
# All rights reserved. Confidential and proprietary.
```

**Constraints**:
- No public distribution of source code
- No sharing of internal implementation details in public docs or logs
- Docstrings should avoid exposing business logic specifics

## Impact on Project Formalism

| Aspect | GPL v3 | MIT / BSD | Apache 2.0 | Proprietary |
|--------|--------|-----------|------------|-------------|
| File header | Full GPL block | SPDX one-liner | Full Apache block | Copyright + confidential |
| LICENSE file at root | Required | Required | Required | Optional |
| NOTICE file | Not required | Not required | Required if exists | Not required |
| Dependency licenses | Must be GPL-compatible | No restriction | No restriction | Verify usage rights |
| Public docstrings | No restriction | No restriction | No restriction | Avoid business logic |
| CHANGELOG | Recommended (changes must be documented) | Optional | Recommended (state changes) | Optional |

## Dependency Compatibility

When adding dependencies, verify license compatibility:

- **GPL v3**: Only GPL-compatible licenses (MIT, BSD, Apache 2.0, LGPL). Reject proprietary or GPL-incompatible deps.
- **MIT / BSD**: Any license is acceptable.
- **Apache 2.0**: Most open-source licenses. Avoid GPLv2-only deps.
- **Proprietary**: Verify each dependency allows commercial/proprietary use.
