# Third-Party Notices — PolyDonky

PolyDonky includes or links against the following third-party software.
Each component is distributed under its own license, reproduced in full below.
The presence of a component in this list does not imply that PolyDonky is
endorsed by its authors.

---

## Shipped Dependencies

These libraries are bundled with or linked into the PolyDonky executable.

---

### CommunityToolkit.Mvvm 8.4.0

**Copyright (c) .NET Foundation and Contributors**
**License: MIT**
<https://github.com/CommunityToolkit/dotnet/blob/main/License.md>

```
MIT License

Copyright (c) .NET Foundation and Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

### DocumentFormat.OpenXml 3.5.1

**Copyright (c) Microsoft Corporation and Contributors**
**License: MIT**
<https://github.com/dotnet/Open-XML-SDK/blob/main/LICENSE>

```
MIT License

Copyright (c) Microsoft Corporation and Contributors.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

### Markdig 0.42.0

**Copyright (c) 2018-2025, Alexandre Mutel**
**License: BSD 2-Clause "Simplified" License**
<https://github.com/xoofx/markdig/blob/master/license.txt>

```
BSD 2-Clause License

Copyright (c) 2018-2025, Alexandre Mutel
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```

---

### .NET 10 Runtime

**Copyright (c) .NET Foundation and Contributors**
**License: MIT**
<https://github.com/dotnet/runtime/blob/main/LICENSE.TXT>

```
MIT License

Copyright (c) .NET Foundation and Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## Test-Only Dependencies

The following packages are used only during development and testing.
They are **not** bundled with or distributed as part of the PolyDonky executable.

| Package | Version | License |
|---------|---------|---------|
| xunit | 2.9.3 | Apache-2.0 |
| xunit.runner.visualstudio | 3.1.4 | Apache-2.0 |
| Microsoft.NET.Test.Sdk | 17.14.1 | MIT |
| coverlet.collector | 6.0.4 | MIT |

---

*This file is generated and maintained manually. Last updated: 2026-05-06.*

---

## HWPX / OWPML Specification References (Documentation Only)

PolyDonky 의 HWPX 코덱(`PolyDonky.Codecs.Hwpx`)은 다음 공개 자료를
참고해 자체 구현했습니다. 아래 자료의 **코드는 PolyDonky 에 직접
포함되거나 재배포되지 않으며**, 명세 이해를 위한 참고 문헌으로만
사용했습니다. HWPX 의 한컴 호환 동작을 정확히 맞출 수 있도록
정보를 공개해 주신 분들께 감사드립니다.

### Specifications

- **한국산업표준 KS X 6101 — OWPML (Open Word-processor Markup Language)**
  - Standardized by KATS (Korean Agency for Technology and Standards) — Korean Government
  - https://standard.go.kr/KSCI/api/std/viewMachine.do?reformNo=03&tmprKsNo=KSX6101&formType=STD
  - National standard for the OWPML document format. HWPX is the file format
    implementation of this standard.

- **HWP/OWPML Format Documentation by Hancom Inc. (한글과컴퓨터)**
  - https://www.hancom.com/support/downloadCenter/hwpOwpml
  - Vendor-published format specification documents and notes.

### Reference Implementations (Documentation Only)

- **`hancom-io/hwpx-owpml-model`** — Apache License 2.0
  - Copyright © 2022 Hancom Inc.
  - https://github.com/hancom-io/hwpx-owpml-model
  - Official OWPML C++ reference model. Used to verify element/attribute
    definitions, child structures, and serialization order for `CPictureType`,
    `CRectangleType`, `CLineType`, `CTableType`, and related shape classes.
  - **No code from this repository is included in or distributed with PolyDonky.**

- **`ai-screams/HwpForge`** — License per repository
  - https://github.com/ai-screams/HwpForge
  - Independent Rust implementation of HWPX read/write. Used as a cross-reference
    for serialization patterns and edge cases.
  - **No code from this repository is included in or distributed with PolyDonky.**

### Community Resources

- **KS X 6101 ↔ Hancom 구현 차이 정리 스프레드시트**
  - https://docs.google.com/spreadsheets/d/1jqXPUVZv1QYcoruJgek2GKYXkhbyaZ68cDjbb1MeyYk/edit
  - Community-maintained document tracking errata in the KS X 6101 standard
    and points where Hancom Office's actual HWPX implementation diverges
    from the standard.
