# BasicAccessories

BasicAccessories is a semi complete
[Microsoft Access](https://products.office.com/en-us/access) file with a
generic VBA code library.

One of the major goal is to be a boilerplate for Microsoft Access applications.

The other one is to provide a set of compatible, tested, generic
[VBA](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications) code
that can directly be used in other Microsoft Office software.
Where possible, the code will compatible with [FreeBASIC](http://freebasic.net/),
[Visual Basic .NET](https://msdn.microsoft.com/en-us/vstudio/hh388573) and
[OpenOffice.org BASIC](https://wiki.documentfoundation.org/Documentation/BASIC_Guide).


# User Guide

## Usage

* Copy `BasicAccessories.accdb.bin` as `YourNewProject.accdb`
* Open `YourNewProject.accdb`, launch Visual Basic Editor with `Alt+F11`
* Make sure that `Immediate Window` is visible.
If it is not visible, it can be displayed by *View* -> *Immediate Window* menu item,
or simply pressing `Ctrl+g`
* Type the following code into the `Immediate Window`
```vbnet
Call mdlLoader.ImportModulesFromDisk()
```

That's it.
Your new project will import all the functions and subroutines for BasicAccessories.
You can then add your own modules to the file.


## Documentation

The documentation of BasicAccessories is compatible with
[Natural Docs](http://www.naturaldocs.org/).
The HTML files can be generated using Natural Docs.


## Structure

BasicAccessories is an almost (it only includes [a module](mdlLoader.bas)
to import other modules) empty Microsoft Access file that is configured to
import the VBA code from the accompanying `.bas` files.

Keeping the VBA code in separate files provides the ability to track
the code changes in [Git](http://git-scm.com/).
It also allows multiple authors to work on the software.
Otherwise, keeping the code in the binary `.accdb` file would disable
tracking the changes on the code.

Please keep in mind that the modules in the `.accdb` files are overwritten
each time the modules are imported.
So, all the permanent changes must be done on the `.bas` files, never on the
`.accdb` file except testing purposes.

Also note that Microsoft Access updates the file `BasicAccessories.accdb` each time it
is opened. So, make sure that you are not committing the file to the repository
if nothing is actually changed. To prevent it,
the file has been distributed as `BasicAccessories.accdb.bin` and `BasicAccessories.accdb`
is already added to [.gitignore](.gitignore).

## How to Import Modules

To load the modules, simply open the `.accdb` file,
and type the following code in the *Immediate Window*:

```vbnet
Call mdlLoader.ImportModulesFromDisk()
```

The output in the *Immediate Window* would resemble this:

```
Call mdlLoader.ImportModulesFromDisk()
Mew module to add : mdlDatabase
Mew module to add : mdlDatabaseTest
Mew module to add : mdlDate
Mew module to add : mdlDateTest
Mew module to add : mdlDrafts
Mew module to add : mdlFiles
Mew module to add : mdlFilesTest
Mew module to add : mdlStrings
Mew module to add : mdlStringsTest
Mew module to add : mdlUnitTestLib
Mew module to add : mdlUnitTestRunner
```

It means that the modules are imported and ready to use.

The list of files to be imported are defined in:
[basicaccessories_modules_to_import.txt](basicaccessories_modules_to_import.txt).
If a new module is added, its name must be defined to this file.


# Developer Guide

## Style Guide
* 4 spaces for indentation, and no tabs.
* Two lines between functions and subroutines.
* At most one conscutive empty lines in functions and subroutines.
* No empty lines before `End Function` and `End Subroutine`.
* No empty lines after `Public Sub ...` and `Public Function ...`
* All the functions and subroutines must have an explicit access modifier such as
`Public Function ...` or `Private Function ...`
* All the functions, subroutines their parameters and returns values must be
[documented](http://www.naturaldocs.org/documenting.html) according to
[Natural Docs](http://www.naturaldocs.org/).
* All the modules should start with:
```vbnet
Option Compare Database
Option Explicit
```
* Each module must have a separate a unit test module.

## Generating Documentation

TODO: Generating documentation. (possibly using Natural Docs)

## Unit Tests

A module that includes a set of very basic unit testing functions and
subroutines, [mdlUnitTestLib.bas](mdlUnitTestLib.bas)
is included with the software.
It provides public subroutines and functions so that other modules
could directly use it.

[mdlUnitTestRunner.bas](mdlUnitTestRunner.bas)
runs all the unit tests.

The following code executes the unit tests:

```vbnet
Call mdlUnitTestRunner.RunAllUnitTests()
```

## How to contribute?

You are welcome to clone this repository, add new functions and subroutines
or improve existing ones in the existing modules,
or add new modules.

`.accdb.bin` file is to be updated rarely, but `.bas` files are very likely
to be modified very often.

You can also contribute to documentation, which also leads to the
improvements of the `.bas` files, since the documentation
(except this very `README.md` file)
is also auto generated.

## How to add a new module?

* Naming first, it should be renamed as `mdlMODULENAMEHERE.bas`
* Make sure it also has test file.
If the name of your module is `MODULENAMEHERE`, the test module should be renamed as
`mdlMODULENAMEHERETest.bas`
* Add both file names into the [basicaccessories_modules_to_import.txt](basicaccessories_modules_to_import.txt)
* Make sure both the files have the required documentation.
* Make sure both the files have the required options:
    ```vbnet
    Option Compare Database
    Option Explicit
    ```
* The test file should have a main sub to call the all the test subroutines in the corresponding file.
* This main module should be called from [mdlUnitTestRunner.bas](mdlUnitTestRunner.bas)


## How to add code to an existing module?

TODO: How to add code to an existing module?

## How to deploy?

TODO: How to deploy

## How to use the modules in another Access file?

* Clone this repository to a folder in your computer.
* Copy your `.accdb` file to to the folder above.
* Open your `.accdb` file.t
* Add a new module `mdlLoader`
* Copy all the contents of the file `mdlLoader.bas` into `mdlLoader` module.
* Switch to `Immediate Window`.f
* Type the following code:
```vbnet
Call mdlLoader.ImportModulesFromDisk()
```
* This will load all the modules to your `.accdb` file.

Note that you will be bound by the [LICENSE](LICENSE).


## Requirements

### Development Requirements

* The file and the function are fully tested with Microsoft Access 2013.
That is the only software you need.

### Usage Requirements

* [Microsoft Access](https://products.office.com/en-us/access) or
[Microsoft Access 2013 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=39358).

# Resources

# Microsoft Access
* [Office 2013 VBA Documentation](https://www.microsoft.com/en-us/download/details.aspx?id=40326)
* [Microsoft Access 2013 Runtime](https://www.microsoft.com/en-us/download/details.aspx?id=39358)

## FreeBASIC
* [FreeBASIC Documentation](http://www.freebasic.net/docs)
* [FreeBASIC Programmer's Guide](http://www.freebasic.net/wiki/wikka.php?wakka=CatPgProgrammer)

## Natural Docs
* [Natural Docs](http://www.naturaldocs.org/)
* [Using Natural Docs with FreeBASIC source code](http://www.freebasic.net/forum/viewtopic.php?f=7&t=11442)
* [Documenting Your Code](http://www.naturaldocs.org/documenting.html)
* [Natural Docs Walkthrough](http://www.naturaldocs.org/documenting/walkthrough.html)
* [Natural Docs Keywords](http://www.naturaldocs.org/keywords.html)

# License

Licensed with 2-clause license ("Simplified BSD License" or "FreeBSD License").
See the [LICENSE](LICENSE) file.


# Legal

All trademarks and registered trademarks are the property of their respective owners.
