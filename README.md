# PyPath - VBA Implementation of Python's os.path

[![VBA](https://img.shields.io/badge/language-VBA-orange.svg)](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications)

## Overview

**PyPath** is a VBA module that provides a direct translation of Python's `os.path` functionality into VBA. This module enables users to manipulate file paths in Microsoft Excel with familiar Python-like methods.

## Features

Provides most of equivalent functions to Python's `os.path`:

- Normalize and convert paths to absolute or relative versions
- Extract directory names, file names, and extensions
- Handle common path operations (join, split, basename, dirname, etc.)
- Check file and directory existence
- Get file metadata (size, creation time, modification time, etc.)
- Expand environment variables and user home paths
- Designed for Windows file system paths

## Installation

### With [ppm](https://github.com/artemdorozhkin/ppm.git)

Run from the Immediate Window:

```vba
ppm "install pypath"
```

### Manually

1. Open **Excel VBA Editor** (`ALT + F11`)
2. Go to **File > Export File...** (`Ctrl + E`)
3. Select the `PyPath.bas` VBA module code
4. Save and start using the functions in your macros

## Functions & Usage

Below are some key functions with usage examples.

### Get Absolute Path

```vba
Dim AbsolutePath As String
AbsolutePath = PyPath.AbsPath("C:\Users\User\..\Documents")
Debug.Print AbsolutePath  ' Output: C:\Users\Documents
```

### Get File Name from Path

```vba
Dim FileName As String
FileName = PyPath.Basename("C:\Users\User\file.txt")
Debug.Print FileName  ' Output: file.txt
```

### Join Path Components

```vba
Dim FullPath As String
FullPath = PyPath.Join("C:\Users", "User", "Documents", "file.txt")
Debug.Print FullPath  ' Output: C:\Users\User\Documents\file.txt
```

### Check if a Path Exists

```vba
Dim Exists As Boolean
Exists = PyPath.Exists("C:\Windows")
Debug.Print Exists  ' Output: True
```

### Get File Size

```vba
Dim FileSize As Long
FileSize = PyPath.GetSize("C:\Users\User\file.txt")
Debug.Print FileSize  ' Output: <file size in bytes>
```

### Normalize Path

```vba
Dim NormalizedPath As String
NormalizedPath = PyPath.NormPath("C:\Users\..\Documents\.")
Debug.Print NormalizedPath  ' Output: C:\Documents
```

## Compatibility

- **Tested Applications**: Excel 2019
- **OS Compatibility**: Windows (not tested on Mac)

## Contributing

Contributions are welcome! If you want to improve the module, submit a pull request or create an issue.

---

Star this repository if you find it useful!
