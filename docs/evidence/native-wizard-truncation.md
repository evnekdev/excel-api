# Function Wizard final-description truncation (#138)

Native parent #95; related #129 and #126. Native sub-issue association #95 → #138
was created through GitHub's sub-issue API.

The maintainer supplied an image from the scratch Function Arguments dialog
for RUST.TEST.HELP during run 20261107T070651Z-5be6130d. It showed the function
help `Acceptance help metadata.`, argument label Value, value/result 17, and
Help on this function. The argument description was `Acceptance argumen`,
missing the final t. This was a user-provided image, not agent screen capture.

Microsoft documents truncation of the last argument-help string and prescribes
an extra empty-string argument-help parameter:
https://learn.microsoft.com/en-us/office/client-developer/excel/known-issues-in-excel-xll-development#argument-description-string-truncation-in-the-function-wizard

CallArguments::registration now appends that counted empty string for functions
with descriptions when the 255-argument call limit leaves a slot. Existing
metadata, signature, HelpTopic position and ownership stay unchanged. At 245
visible descriptions the prior 255-slot call is retained without a sentinel;
that boundary's host truncation limitation remains, explicitly documented.
Regression coverage checks 0, 1, 244 and 245 descriptions, valid live counted
string roots, HelpTopic slot and complete final description. Existing exact
position tests now also assert the trailing empty string for absent/supplied
HelpTopic. The initial full test run failed its old 11-argument expectation;
after updating that expected wire contract, the full workspace suite passed.

Corrected native artifact acceptance-wizard-fixed.xll SHA-256:
1342cc7e5f5e62897f894c141be586212ecc9d63a36171de0a73fe3f45fcaa11.
Built from the review working tree based on fabda4d plus the #138 source fix
and cancellation fixtures, not an assertion of historical commit provenance.
Native corrected text observation is pending at this evidence revision.

Native child-window text, UI Automation Name/HelpText and MSAA traversal did
not expose both expected descriptions; automated HELP-WIZARD remains BLOCKED
rather than silently accepting manual evidence as an automated assertion.
The empty-sentinel fix does not promise Excel can open the external help URL.
No help URL was opened, Office setting changed, x86 test or CI run performed.
