These examples automate solving each of the optimization problems in SOLVSAMP.XLS, which is distributed with MS Office Excel and can be found programatically in:
```vba
Application.LibraryPath & "\..\SAMPLES\SOLVSAMP.XLS"
```
... which on many systems resolves to:
```
C:\Program Files\Microsoft Office\root\Office16\SAMPLES\SOLVSAMP.XLS
```

Import these test modules into the sample workbook above, load the [vba source](https://github.com/GCuser99/SolverWrapper/tree/main/src/vba) or set a reference to the [SolverWrapper DLL]( https://github.com/GCuser99/SolverWrapper/tree/main/dist) and then save SOLVSAMP.XLS to SOLVSAMP.XLSM.
