OfficeExtractor
===============

Extracts embedded OLE objects from Word, Excel and PowerPoint files without needing the original programs.

- 2014-06-11 Version 1.2
  - Added support for Open Office files; .ODT, .ODS and .ODP


- 2014-06-11 Version 1.1
  - Moved all the CompoundFileStorage to another project and made it available through a nuget package

- 2014-06-10 Version 1.0

  - Extracts embedded files from binary office files (Office 97 - 2003)
  - Extracts embedded files from Office Open XML files (Office 2007 - 2013)
  - Automaticly sets hidden workbooks in Excel files visible
  - Will detect if the files are password protected
  - Unit tests for the most common used file types
  - 100% native .NET code, no PINVOKE

Note for contributors
=====================

Please create your pull requests to target the "develop" branch. "Master" is only for released code. Thank you.

License
=======

Copyright 2013-2014 Kees van Spelde.

Licensed under the The Code Project Open License (CPOL) 1.02; you may not use this software except in compliance with the License. You may obtain a copy of the License at:

http://www.codeproject.com/info/cpol10.aspx

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
Core Team

    Sicos1977 (Kees van Spelde)
