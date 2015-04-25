DispHelper
==========

Origin
------

This is a fork of xtmouse's excellent [DispHelper](http://disphelper.sourceforge.net/)

* Original: http://disphelper.sourceforge.net/
  * [Readme](http://disphelper.sourceforge.net/readme.htm) | [CVS](http://disphelper.cvs.sourceforge.net/viewvc/disphelper/disphelper/) | [Project page](https://sourceforge.net/p/disphelper/wiki/Home/) | [License](http://opensource.org/licenses/bsd-license.php) | [Forum](http://sourceforge.net/forum/?group_id=111558) | [Tracker](http://sourceforge.net/tracker/?group_id=111558&atid=659678)
  
I've decided to publish my modifications to it, if they might be useful to someone. This is
old work which was done back in 2009.

Description
-----------
DispHelper is a COM helper library that can be used in C++ or even plain C. No
MFC or ATL required! It allows you to call COM objects with a printf style syntax. 
Included with DispHelper are over 20 samples demonstrating using COM objects from ADO to WMI

License
-------

Original work by xtmouse was licensed under [BSD license](http://opensource.org/licenses/bsd-license.php)

Disclaimer
----------

DispHelper has been modified to better serve our project :
* [Captain-bol](https://gitlab.isb-sib.ch/itopolsk/captain-bol/tree/master) - A c++ tool suite for mass spectrometry identification 
  * old [google code](https://code.google.com/p/captain-bol/) page

Modifications on DispHelper were done while working at [GeneBio SA](http://www.genebio.com)

Development has been paid by the [Swiss Institute of Bioinformatics](http://www.isb-sib.ch)

Big thanks to xtmouse for the original code : your work has been really helpful.


Modification
------------

* [patch1](https://gist.github.com/DrYak/81f73fff0d572130ff9d) - patch for single source
* [patch2](https://gist.github.com/DrYak/532c10720065d18e27bd) - patch for split source

Our main project, [Captain-bol](https://gitlab.isb-sib.ch/itopolsk/captain-bol) was  a [proteomics](https://en.wikipedia.org/wiki/Proteomics) software, which thus needs to manipulate massive amount of data (and thus ran better on 64 bits platforms). In order to import data from foreign prorietary formats, it needed to be able to call into proprietary API provided as closed source [OLE/COM objects](https://en.wikipedia.org/wiki/Component_Object_Model). Lastly, Linux is our preferred platform.

To better suit our needs, a few modification have been added to DispHelper for supporting :
* better Wine under Linux
* 64 bits platforms
* call by ref



### Wine on Linux

Small fixes were added to better support Wine under Linux. 
With this it is possible to create native Linux apps which call into OLE/COM object in .DLL

To compile you need to add the following parameters:
* `-lole32 -loleaut32 -luuid`
(see [Makefile.am](https://gitlab.isb-sib.ch/itopolsk/captain-bol/blob/master/xenobol/src/Makefile.am) for an example)

### 64 bits platforms

Support for 64 bits platforms has been added

> **WARNING** keep in mind that Windows is a [LLP64](https://en.wikipedia.org/wiki/64-bit_computing#64-bit_data_models) platform whereas Linux is [LP64](https://en.wikipedia.org/wiki/64-bit_computing#64-bit_data_models) like every other unix

Whereas xtmouse's disphelper only support simple single caracter `%` formats reminiscent of `printf`'s format string (see [Format Identifiers](http://disphelper.sourceforge.net/readme.htm)), we introduce lenghts specifier ( `h`, `l`, `ll`, `L` ) similar to those used in [C99](https://en.wikipedia.org/wiki/Printf_format_string#Format_placeholders) :

| signed  | unsigned | bit width | signed    | unsigned           | comment |
|---------|----------|-----------|-----------|--------------------|---------|
| `%d`    | `%u`     | 32 bits   |`int`      |`unsigned`          |
| `%ld`   | `%lu`    |__32 bits__|`long`     |`long unsigned`     | remember that windows is __LLP64__
| `%lld`  | `%llu`   | 64 bits   |`long long`|`long long unsigned`| correct way to ask for 64 bits

* **Note** as with all LP64 and LLP64 platforms, the default type is 32 bits long, and thus shorter types parameters (`%hd`, `%hu`, etc.) are all automatically promoted to 32 bits.
* On the other hand, when asking for it (`%lld`, `%llu`) 64 bits parameters must be consumed from the parameter stack.

Same for floating point :

| format | bit width | type        | comment |
|--------|-----------|-------------|---------|
| `%e`   | 64 bits   |`double`     | **WARNING** in DispHelper `%f` means `FILETIME *`, don't use it
| `%Le`  | (varying) |`long double`| (on Intel, it's the internal non normalized 80 bits format,<br />with other it might by 128 bits)

* **Note** the default floating point is 64 bits double precision, simple precision 32 bits are promoted automatically
* OLE/COM itself doesn't have support for Intel's 80bits _extended precision_ nor for 128bits quad precision. Internally they are down converted to 64bits double precision, but still consume the correct amount of data from the parameter stack

Remark: due to the way they are internally handled, `ll` and `L` can be used interchangeably, even if that doesn't follow exactly the C99 standard `printf` format string.

### Call by ref

xtmouse's original DispHelper could only call methods by value. Calling by referrence did require using `%v` and manually constructing [VARIANT](https://msdn.microsoft.com/en-us/library/cc237865.aspx) to be used in the message (as with anything else not supported by the current interface).

To simplify this work a new interface has been introduced.
* it uses `&`, inspired by C++'s call by reference
* format identifiers then follows the `scanf` convention for [format string](http://linux.die.net/man/3/scanf)
* parameters are pointer to the corresponding precise types (as with `scanf`)

| signed | unsigned | bit width | signed      | unsigned             | comment |
|--------|----------|-----------|-------------|----------------------|---------|
| `%&hhd`| `%&hhu`  | 8 bits    |`char *`     |`unsigned char *`     | currently disabled as `V_I1REF`/`V_U1REF` macros might be missing
| `%&hd` | `%&hu`   | 16 bits   |`short *`    |`short int * `        |
| `%&d`  | `%&u`    | 32 bits   |`int *`      |`unsigned *`          |
| `%&ld` | `%&lu`   |__32 bits__|`long *`     |`long unsigned *`     | remember that windows is __LLP64__
| `%&lld`| `%&llu`  | 64 bits   |`long long *`|`long long unsigned *`| correct way to ask for a pointer to 64 bits

As for floating point :

| format | bit width | pointer  | comment |
|--------|-----------|----------|---------|
| `%&e`  | 32 bits   |`float *` |__WARNING__ follows `scanf` convention
| `%&le` | 64 bits   |`double *`|

* **Note** we follow `scanf`'s convention that floating point types are simple precision `float *` by default (unlike `printf` which promotes everything to 64bits double precision `double`)
* there's no way to cram support for extended/quad precision floats `long double *` in a compatible way. VARIANT simply don't support that type.

See examples in :
* [arthur.cpp](https://gitlab.isb-sib.ch/itopolsk/captain-bol/blob/master/xenobol/src/arthur.cpp)
* [analyst.cpp](https://gitlab.isb-sib.ch/itopolsk/captain-bol/blob/master/xenobol/src/analyst.cpp)

## Limitations

Currently, only the internal function [`ExtractArgument`](https://github.com/DrYak/disphelper/blob/master/single_file_source/disphelper.c#L589) which handles manipulation of method call parameters has been patched.

The function [`dhGetValue`](https://github.com/DrYak/disphelper/blob/master/single_file_source/disphelper.c#L295) which handles retrieving the return value with the user friendly format string and pointer hasn't been patched yet, and thus doesn't handle the `l`,`ll`,etc. lenght specifier.

It should be possible to replicate the work, but **beware** that currently, DispHelper API defines `%e` as a `double *` (like `printf` does), not as a `float *` (as both `scanf` and our patch does). Changing this behaviour can potentially cause legacy code breakage !!!

## BSTR

This isn't part of the modification done on DispHelper, but might comme handy when handling OLE/COM objects :

* [comutil.h](https://gitlab.isb-sib.ch/itopolsk/captain-bol/blob/master/xenobol/src/comutil.h) - for handling BSTR (BASIC strings)
