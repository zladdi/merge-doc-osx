# merge-doc-osx

A port of TortoiseSVN/TortoiseGit merge-doc.js to Objective-C on Mac OSX using ScriptingBridge.
For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/merge-doc.js.

The code in is distributed under the GNU General Public License. 

## Prerequisites

Microsoft Word needs to be installed.

## Usage 

`merge-doc-osx <absolute-path-to-merged.doc> <absolute-path-to-theirs.doc> <absolute-path-to-mine.doc> <absolute-path-to-base.doc>`

## Build instructions

Open and build the project in XCode. The header file `Word.h` was initially generated
using the follwing command (see also this documentation from [Apple](https://developer.apple.com/library/content/documentation/Cocoa/Conceptual/ScriptingBridgeConcepts/UsingScriptingBridge/UsingScriptingBridge.html))

`sdef /Applications/Microsoft\ Word.app | sdp -fh --basename Word`

The resulting file was adapted so as to be free of compile errors (but not free of warnings).

## Known issues

* Relative paths to documents are not supported.
* Newer versions of Microsoft Office apps are sandboxed and do not allow modifying the document in-memory if the app does not have write access to the underlying file. Thus, `merge-doc-osx` temporarily saves a copy of comparison results to a folder where the app has write access to (`~/Library/Group Containers/UBF8T346G9.Office`). Normally, `merge-doc-osx` deletes the documents containing the comparison results automatically as soon as possible. However, it might happen that the program exits unexpectedly without having removed the documents. So, you might want to routinely check for stale documents in that folder.
* Newer versions of Microsoft Office apps are sandboxed. This leads to the annoying
"Grant File Access" dialog to pop up for each of the documents involved in the merge in cases
where Word does not have permission to access the respective file already.

## Future work

Create a formula for [`brew`](https://github.com/Homebrew)
