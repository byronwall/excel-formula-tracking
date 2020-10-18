# Excel Formula Tracker

This add-in provides a number of functions to walk through the calculation chain in Excel.  The goal of these functions is to allow you to audit and inspect how a given cell arrives at its result.  The calculation chain is walked  until all preceding cells are resolved to constants.

When developed, this tool was used to convert a series of complicated formulas into a set of commands which could be added to a different program's source code.  The goal was to provide a simpler interface for maintaining those calculations (i.e. Excel) without having an expensive translation step into the other system (i.e. LabView).

Various versions of the function tracking were put together depending on the outcome.  One version stopped tracking cells once a named range was met.  This allowed the resulting command to refer to already existing variable names on the LabView side.

In addition to being a tool for tracking formulas, this was also a test bed for general Excel formula parsing and tracking.

## Screenshots

Example shows a couple of cells that refer to each other.  The tool will trace the related calls and generate the final expression that defines the result.  There are several options which are not used which help guide how "deep" into the calculation tree to process.

![Tracer inputs](/docs/input.png)

Result of the call to `GetFormulaOptionsFull` which traces the calc.  Here the final result is `B2*B3*(20 + B2*B3)` which rolls up all the related calls.

![Tracer outputs](/docs/output.png)