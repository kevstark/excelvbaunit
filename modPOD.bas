Attribute VB_Name = "modPOD"
Option Explicit
'This is written for ExcelPOD, and, if processed properly, will produce HTML documentation.
Rem title VbaUnit version 16 - Unit Test harness for Excel VBA - User Documentation
Rem title VbaUnit version 16 - Unit Test harness for Excel VBA - Technical Documentation
Rem order 1
Rem docs 2
Rem doc 1 2
Rem
Rem =head1 Introduction
Rem
Rem =head2 Overview (by MH)
Rem
Rem A unit test framework similar to nunit (L<http://www.nunit.org/>) and
Rem junit (L<http://www.junit.org/>), but for Excel VBA code.
Rem It is similar to the existing VbaUnit project in sourceforge but:
Rem
Rem =over
Rem
Rem =item 1. There is no self-modifying code, so is easier to write tests to test itself.
Rem
Rem =item 2. The code lives in an Excel addin so you do not need to copy the code into each
Rem project to be tested (obviously this makes it Excel specific).
Rem
Rem =item 3. The original sourceforge VbaUnit requires you to call a Prep function prior to
Rem running your tests and remembering to call Prep when you add/remove testing functions. There
Rem is no need to do this with this framework - just call xRun to execute your tests.
Rem
Rem =back
Rem
Rem =head2 To do
Rem
Rem =over
Rem
Rem =item * Make use of Apps Hungarian notation consistent throughout.
Rem
Rem =item * Change coverage analysis to cope with multiple implementations of class modules.
Rem
Rem =item * Write documentation on how to test worksheet formulae (including TDD).
Rem
Rem =item * Write a GUI.
Rem
Rem =item * Change the ErrorTrap constants back to their usual status as variables, including
Rem the options either on the GUI or on a sheet (usually very hidden). This will enable
Rem unattended, automated testing that can detect errors from the log file.
Rem
Rem =item * Add class modules to enable test results to be logged to a file as another necessity
Rem for unattended automated testing.
Rem
Rem =back
Rem
Rem =head2 Change Log
Rem
Rem =head3 Changes from version r15
Rem
Rem =over
Rem
Rem =item * Fixed a bug (divide by zero) in AssertEqual that was introduced in r14.
Rem
Rem =back
Rem
Rem =head3 Changes from version r14
Rem
Rem r15 was an additional version number generated automatically when a wiki page was added.
Rem
Rem =head3 Changes from version r13
Rem
Rem =over
Rem
Rem =item * Wrote documentation in ExcelPOD.
Rem
Rem =item * Added comprehensive error trapping.
Rem
Rem =item * Refactored certain procedures to avoid looping twice by using ReDim Preserve.
Rem Certain other procedures were refactored out in consequence as they were no longer useful.
Rem
Rem =item * Fixed issue 2 (L<http://code.google.com/p/excelvbaunit/issues/detail?id=2>)
Rem to allow test modules with no tests. This is important because users may already have
Rem modules called *Tester to test for other things. These would automatically be treated by
Rem xRun as modules expected to contain tests. It also allows projects without test modules.
Rem
Rem =item * Fixed issue 1 (L<http://code.google.com/p/excelvbaunit/issues/detail?id=1>). This
Rem now produces a messagebox if the project can't be found.
Rem
Rem =item * Fixed a bug in TestRunnerTester that caused it to report zero failures and zero
Rem successes, not only for itself but for all test modules run subsequently.
Rem
Rem =item * Changed restrictions on subs that contain tests to allow functions and subs that
Rem are not explicitly declared as Public, but still excluding Private and Friend.
Rem
Rem =item * Added a retrofitter to create a new Tester module and Test stubs for any module
Rem that does not already have one.
Rem
Rem =item * Added a coverage analyst.
Rem
Rem =item * Warns when testing if there is a SetUp procedure but no TearDown.
Rem
Rem =item * Refactored tests that look for modules in specific orders to take them in any order.
Rem
Rem =item * Changed AssertEqual to cope with floating point differences.
Rem
Rem =item * Retrofitted tests to most of the procedures in Assert.
Rem
Rem =item * Changed tests for modules so that Excel did not need to store them in any specific
Rem order.
Rem
Rem =item * Retrofitted a new MainTester as the one needed by the tests was not in the
Rem repository.
Rem
Rem =back
Rem
Rem =head3 Changes in previous versions
Rem
Rem See L<http://code.google.com/p/excelvbaunit/>.
Rem
Rem =head2 Authors
Rem
Rem The original author (and writer of most of the code) was Matt Helliwell (MH). The error
Rem trapping, coverage and retrofit code and most of the documentation were written by John
Rem Davies (JHD).
Rem
Rem =head2 Copyright
Rem
Rem L<http://www.gnu.org/licenses/lgpl.html>
Rem
Rem doc 1
Rem =head1 Usage
Rem
Rem =head2 First write your tests!
Rem
Rem If you are new to testing and have a project to which you would like to add tests,
Rem the lazy way is to run RetroFit on all modules that contain testable code
Rem and then use this skeleton to write your tests.
Rem
Rem The "proper" way to use tests is in Test Driven Development (TDD). Searching will give a lot
Rem of results as this is a commonly used technique. TDD suggests that the following approach
Rem should be followed:
Rem
Rem =over
Rem
Rem =item 1 Write a test.
Rem
Rem =item 2 Run it and check that it fails.
Rem
Rem =item 3 Now - only now - write the code to make it pass.
Rem
Rem =back
Rem
Rem Repeat this sequence until you have something that does what you want.
Rem
Rem However, I would add a step between 1 and 2 above. Write the test to call the new code
Rem you are going to write and make sure it doesn't compile. This step makes sure you are not
Rem by accident repeating a procedure name.
Rem
Rem VbaUnit is written as an add-in. However, it is necessary to put a reference to it in the
Rem project you wish to test.
Rem
Rem Tests should be written in standard code modules of their own. The modules should
Rem end with "Tester". The tests should begin with "Test". VbaUnit will then be able to find
Rem them automatically.
Rem
Rem The project should have its own, distinct name. Excel defaults to calling every project
Rem "VbaProject". It is good standard practice to change this, but it is almost essential when
Rem using VbaUnit, as the project name needs to be given, and having two projects with the
Rem same name will confuse VbaUnit. This is very likely if "VbaProject" is retained for
Rem everything.
Rem
Rem Individual tests should be written in subs or functions with names starting with "Test".
Rem The organisation of tests into subs is at the whim of the writer, but the following scheme
Rem is advised:
Rem
Rem =over
Rem
Rem =item * Each code module has its own "Tester" module.
Rem
Rem =item * Each sub or function has its own "Test" sub or function.
Rem
Rem =back
Rem
Rem There is no obvious way to write tests for the class module procedures "Property Get",
Rem "Property Let" and "Property Set". If a project requires more than trivial code in
Rem such a procedure, it should be abstracted to a function or sub called from the
Rem procedure. This function or sub can then be tested in the usual way.
Rem
Rem Variables that are declared at module level are hard to test. Either they have to be
Rem declared Public, which would make them accessible to the Tester module but which is a
Rem Bad Thing(tm), or the test routines have to be written in the live module. Not only does
Rem this make test routines dangerous - they might be invoked by accident by a maintenance
Rem programmer - but they cannot be run automatically by xRun, nor can the coverage calculator
Rem identify them properly. The "proper" way to do this is by means of multiple implementations
Rem of class modules. This project demonstrates how this is achieved in the "TestResultsManager"
Rem class modules and the "TestResultsManagerTester" standard code module. However, it is not
Rem for the faint hearted, requiring a detailed understanding of how class modules implement
Rem overloading. If this sounds too technical for you, it's best to accept that you can't test
Rem assignments to module level variables and that coverage will be less than 100%. You can
Rem minimise the risks of this by doing as little as possible in routines that modify module
Rem level variables and moving everything else to procedures that don't modify module level
Rem variables. You can also fool the coverage calculator by ignoring the advice about
Rem segregating code and letting the coverage calculator report that all procedures have test
Rem procedures. This might get past a pointy haired boss, but it won't mean that you are testing
Rem your code adequately. This merely shows up one of the deficiencies in the coverage
Rem calculator. While it can determine whether a test procedure exists, it can't determine if
Rem it does the job properly.
Rem
Rem There are technical reasons why it is usually better to write functions than subs. This
Rem project is itself an example of one of the benefits. Every procedure is wrapped in error
Rem trapping, meaning that an error can pass a failure code back to the calling procedure.
Rem This, in turn, raises an error which is passed to its calling procedure and so on until
Rem the code terminates. This means that there is no code that shows the user the IDE and also
Rem that it is possible to enable a "Full Trace" mode that shows how a procedure that throws
Rem an error was called. This can make debugging much easier.
Rem
Rem However, the advantages of writing everything as functions are relatively minor and it is
Rem not the purpose of this system to compel a certain coding style. While there are no
Rem unnecessary subs in this system, that does not mean that no user will ever write another
Rem sub. The system can handle and test subs and functions with equal ease.
Rem
Rem The three test evaluators are
Rem AssertTrue, AssertFalse and AssertEqual. The first two test for conditions, while the last
Rem compares strings and numbers. Be aware of the standard issues in digital computing when
Rem using AssertEqual with floating point numbers. This is discussed in more detail in the
Rem section on L<AssertEqual|/AssertEqual(Expected as variant, Actual as variant)>.
Rem
Rem =head3 If you need to change the spreadsheet before your tests can run ...
Rem
Rem If a sub called "SetUp" exists, it will be invoked by the test process before any tests are
Rem run. This will
Rem happen regardless of where such a sub appears in the module, but it must be in the same
Rem module as the tests themselves. If your mother doesn't live in the spreadsheet, you can tidy
Rem up after yourself by means of another sub called "TearDown".
Rem
Rem =head2 Then run them.
Rem
Rem Running the tests is easy. From the IDE (the interface where you write your code), hit
Rem {Ctrl G} to get to the Immediate pane and type C<xrun projectname>, replacing
Rem C<projectname> with the name of your project. This will run all the tests in the project
Rem automagically.
Rem
Rem Tests must be run from the Immediate pane. The command "xRun <project name>" will run all
Rem tests in the named project. Optionally, you can name the module containing the tests you
Rem want to run. This is useful if you have multiple modules containing tests, but want to run
Rem only some of them. There is no provision for running only certain tests within a module.
Rem
Rem Results are logged to the Immediate pane. It is planned to enable logging to a file - see
Rem L<To Do>.
Rem
Rem doc 2
Rem =head1 Technical Documentation
Rem
Rem C<Option Private Module>, while generally recommended, has not been used as the add-in
Rem needs to be accessible to modules in other files. It I<has> been used in "*Tester" and
Rem "Dummy*" modules that should not be used by other files. It cannot be used in class
Rem modules.
Rem
Rem =head2 Hints for Developers
Rem
Rem Before starting development, type C<xrun> in the Immediate pane. This will run all
Rem tests currently in the system. It makes sense to note the number of tests at the end of
Rem the run. Then, whenever you make a change, I<no matter how slight>, run all the tests and
Rem check that everything passes. If you make changes that break something else, you will
Rem find that out as early as possible and be able to backtrack or compensate depending on
Rem what the situation requires.
Rem
Rem =head2 Named Ranges
Rem
Rem There are no named ranges in the add-in.
Rem
Rem =head2
Rem sheetname
Rem
Rem This module contains nothing but POD.
Rem
