---
layout: post
title: "Extracting userform field values from VBA maldocs"
categories: analysis
---

I came across a [maldoc] recently which uses strings stored as values of various OLE form fields to build an executable file. It also constructs various commands to drop and run the executable in the same way. The maldoc is [probably] related to [Lazarus Group], but this post is not about attribution.

A total of 70 form fields stored as Caption, Tag, Text or ControlTipText values of various control elements (like TextBox, CommandButton, ScrollBar etc.) embedded across four parent Form controls are present in the maldoc.
These values are concatentated (at times, after an StrReverse() operation) in the VBA macro to build a hex string of the executable. Similarly, the macro also builds multiple text strings which are passed as parameters to functions responsible for creating the executable file on disk and executing it.

![Macro code showing references to form field values](/assets/post_images/2022-06-28-extracting-vba-userform-field-values/macro_snippet.png)

Extracting these strings is, of course, possible by running the macro code in the Office VBA debugger and setting breakpoints at appropriate locations. However, I endeavoured to extract the strings "statically".

Considering that VBA maldocs are very common initial infection vector for all sorts of malware, I expected this task to be simple enough and hoped to easily find tools with capability to extract form field values from embedded controls.
It was surprising, therefore, when it took me considerable effort to do this properly.

This post hopes to be useful to others attempting something similar.

### Search for a solution
[Didier Stevens], author of multiple excellent tools including [oledump], in an old Twitter thread [suggested] extracting the values using _strings_ on the OLE stream after one has located the name of the corresponding control element.
This does not seem particularly enticing to do (or automate) 70 times, and will actually not work in certain cases mentioned later in this post where the name of the control element and the values of some of its form fields are stored in separate streams.

In the same Twitter thread, [Philippe Lagadec](https://twitter.com/decalage2) [pointed] to a component of their [oletools] project named _oleform.py_ which contains the code to parse the OLE form streams and extract form field values.

Structures and specifications related to OLE form streams are described by Microsoft in [[MS-OFORMS]:
Office Forms Binary File Formats](https://interoperability.blob.core.windows.net/files/MS-OFORMS/[MS-OFORMS].pdf).

The function `extract_OleFormVariables()` of _oleform_ takes an OLE file and a stream directory as parameters and returns a list of dictionaries containing _{field_name: value}_ pairs for each form control in the stream.
The _oletools_ project however does not contain any interface to use `extract_OleFormVariables()` directly on an input file. A call to the function exists in `extract_form_strings_extended()` function of _olevba.py,_ however this function itself is not used anywhere.

I wrote a simple wrapper to call `extract_OleFormVariables()` for the maldoc and see if it works.

![Wrapper code to call extract_OleFormVariables()](/assets/post_images/2022-06-28-extracting-vba-userform-field-values/initial_wrapper.png)

### Encountering a bug

Many of the output values were prepended with 1-3 bytes of seemingly junk data. These bytes are not present when the form field value is viewed in Word or LibreOffice Writer. An equal number of bytes were also missing at the end of the value in these cases.

![Difference in values from oleform and LibreOffice Writer](/assets/post_images/2022-06-28-extracting-vba-userform-field-values/value_difference.png)

This indicated presence of a bug in _oleform_, which I decided to try locating and fixing. After some debugging and going through MS-OFORMS, I found the bug.
The properties Name, Caption, Tag, Text, ControlTipText etc. of various controls are all of type String (specifically an [fmString]). These strings are padded to multiples of four bytes in the OLE stream. _oleform_ did not consider this padding in String values while parsing the stream.

The bug is understandable as the fact that the strings are padded is not mentioned anywhere in the description of any controls in MS-OFORMS. Rather, it is only mentioned briefly at the end of [section 2.1.1.2.4 - Padding and Alignment]:

> _... Property values that are strings are padded to a length that is divisible by 4 ..._

After adding code in _oleform_ that takes care of padding while reading string values, I ran my wrapper again.

This took care of those extra junk bytes and it seemed that my work was done. Not quite!

### Requiring an extra handler

I soon discovered that some of the controls had a Caption property (called Label in LibreOffice Writer) which was not extracted by _oleform_.

![Missing caption/label value for a Frame control](/assets/post_images/2022-06-28-extracting-vba-userform-field-values/missing_caption.png)

All these controls had [ClsidCacheIndex value of 14] indicating it is a Frame control - i.e. a control embedded in a parent Form while having child controls of its own. ([section 2.1.2.2.2 - Embedded Parents])  
In such cases, the OLE stream directory contains a child "i" stream which contains the missing Caption property. (and can contain additional properties)

I added a function `consume_EmbeddedFormControl()` in _oleform_ to handle this case for controls having ClsidCacheIndex value 14. The function only extracts the Caption property as that was enough for this maldoc.

Finally, I had all 70 form fields needed to construct the executable Hex string and other text strings.

### Finally done

I added some hacky code to my wrapper to convert VBA expressions to Python expressions using Regex and evaluate them using Python `eval()`. The final code is available [here].

The file created from the hex string by the macro is [SHA256: 83388741cb6e6ee7341ae00cb9ab92c8a0132d92307473a4c08e678153a27cef]. 

_Sidenote: The above file dropped on the disk by the macro contains one extra null byte than the file which one would get from the concatenated hex string. This difference comes from use of `Redim` with size `len(hexstring)/2` in the macro without using [Option Base 1]. Due to this, the array index in `Redim` statement starts from 0 instead of 1, leading to an extra null byte in the end._

Other constructed text strings by the macro are shown below:

![Strings constructed by the macro](/assets/post_images/2022-06-28-extracting-vba-userform-field-values/constructed_strings.png)

The file is created in _%TEMP%_ folder as _1Vqar5tGI51.sak._ It is then moved to _Startup_ folder and renamed as _WHealthScanner.exe._ Finally, the file is executed.

Executable file is a basic malware with capability to run obtained commands and send their output to a hardcoded C2 server, and download and install additional executables. Interestingly, the malware is written in [PureBasic] which is quite rare. I plan to share brief details on the operation of the malware in an upcoming post.

---
---
[maldoc]: https://www.virustotal.com/gui/file/303bc0f4742c61166d05f7a14a25b3c118fa3ba04298b8370071b4ed19f1a987/details
[probably]: https://xkcd.com/285/
[Lazarus Group]: https://attack.mitre.org/groups/G0032/
[Didier Stevens]: https://twitter.com/DidierStevens
[oledump]: https://blog.didierstevens.com/programs/oledump-py/
[suggested]: https://twitter.com/DidierStevens/status/1225378445040988161
[pointed]: https://twitter.com/decalage2/status/1225687663123939329
[oletools]: https://github.com/decalage2/oletools
[fmString]: https://interoperability.blob.core.windows.net/files/MS-OFORMS/[MS-OFORMS].pdf#%5B%7B%22num%22%3A309%2C%22gen%22%3A0%7D%2C%7B%22name%22%3A%22XYZ%22%7D%2C69%2C473%2C0%5D
[section 2.1.1.2.4 - Padding and Alignment]: https://interoperability.blob.core.windows.net/files/MS-OFORMS/[MS-OFORMS].pdf#%5B%7B%22num%22%3A117%2C%22gen%22%3A0%7D%2C%7B%22name%22%3A%22XYZ%22%7D%2C69%2C641%2C0%5D
[ClsidCacheIndex value of 14]: https://interoperability.blob.core.windows.net/files/MS-OFORMS/[MS-OFORMS].pdf#%5B%7B%22num%22%3A289%2C%22gen%22%3A0%7D%2C%7B%22name%22%3A%22XYZ%22%7D%2C69%2C242%2C0%5D
[section 2.1.2.2.2 - Embedded Parents]: https://interoperability.blob.core.windows.net/files/MS-OFORMS/[MS-OFORMS].pdf#%5B%7B%22num%22%3A123%2C%22gen%22%3A0%7D%2C%7B%22name%22%3A%22XYZ%22%7D%2C69%2C401%2C0%5D
[here]: https://gist.github.com/nvdp01/b557202a49be950ce699ddae8d94249b
[SHA256: 83388741cb6e6ee7341ae00cb9ab92c8a0132d92307473a4c08e678153a27cef]: https://www.virustotal.com/gui/file/83388741cb6e6ee7341ae00cb9ab92c8a0132d92307473a4c08e678153a27cef/details
[Option Base 1]: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/option-base-statement
[PureBasic]: https://en.wikipedia.org/wiki/PureBasic