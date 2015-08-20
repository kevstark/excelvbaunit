A unit test framework similar to nunit and junit for but Excel VBA code. It is similar to the existing VbaUnit project in source forge but
1. Does not self-modifying code so is easier to write test to test itself.
2. The code lives in an Excel addin so you do not need to copy the code into each project to be tested (obviously this makes it Excel specific)
3. The original source forge VbaUnit requires you to call a Prep function prior to running your tests and remembering to call Prep when you add/remove testing functions. There is no need to do this with this framework - just call xRun to execute your tests.

You can direct any queries etc. to me at matt AT helliwell DOT me DOT uk

NEWS
22 Dec 2006
Larger example now completed

20 Dec 2006
Started writing some example spreadsheets in lieu of documentation. Simple example is complete, larger example is still in development.

9 Dec 2006
Added prototype gui interface so the tests can be run from a gui rather than the immediate window in the VBA editor.

8 Dec 2006
Added support for fixture level setup/teardown functions
The xRun routine now allows you to run individual test fixtures rather than all tests.

7 Dec 2006
Added support for SetUp/TearDown functions in test cases
This software is now licensed to (and being used by) Barclays Capital

10th Sept 2006 - Just about got everything under test. Not as pretty as it might be but at least now I'm safe to start changing it.

1st Sept 2006 - Project is pre-alpha. I've written enough code to be able to start self-tests





