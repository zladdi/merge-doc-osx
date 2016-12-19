//
//  main.m
//  merge-doc-osx
//
//  A port of TortoiseSVN/TortoiseGit merge-doc.js to Objective-C on Mac OSX using ScriptingBridge.
//  For the source script, see https://github.com/TortoiseGit/TortoiseGit/blob/master/contrib/diff-scripts/merge-doc.js.
//
//  Word.h header had to be adapted after having been generated. This is a known issue, see e.g.
//    http://stackoverflow.com/questions/15338454/scripting-bridge-and-generate-microsoft-word-header-file
//
//  The Objective-C code was partly inspired by this forum post:
//    https://discussions.apple.com/thread/2623068
//
//  This file is distributed under the GNU General Public License.
//
//  Author: Zlatko Franjcic

#import <Foundation/Foundation.h>
#import "Word.h"

NSString* executeCompare(WordApplication *word, NSString *sBaseDoc, NSString *sOtherDoc, NSString *sOtherDocAuthor)
{
    WordDocument *baseDoc, *otherDoc;
    NSString *sTargetDoc;
    
    uint vOffice2013 = 15;
    WordWdCompareTarget wdCompareTarget = ([[word version] intValue] < vOffice2013 ? WordWdCompareTargetCompareTargetSelected : WordWdCompareTargetCompareTargetNew);
    
    // No 'activate' method -> comment code
    //[baseDoc activate]; //required otherwise it compares the wrong docs !!!
    // We cannot activate the document, so we open it, which should activate it
    baseDoc = [word open:nil fileName:sBaseDoc confirmConversions:YES readOnly:NO addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];
    
    [baseDoc comparePath:sOtherDoc authorName:sOtherDocAuthor target:wdCompareTarget detectFormatChanges:YES ignoreAllComparisonWarnings:YES addToRecentFiles:NO];
    
    if ([[word version] intValue] < vOffice2013)
    {
        sTargetDoc = sOtherDoc;
    }
    else
    {
        // Due to sandboxing, we are going to save the new document to a temporary location that MS word has access to
        // This approach was inspired by: http://www.rondebruin.nl/mac/mac034.htm
        // Group container URL: https://developer.apple.com/library/content/documentation/Security/Conceptual/AppSandboxDesignGuide/AppSandboxInDepth/AppSandboxInDepth.html
        // The temporary file handling part was inspired by: http://nshipster.com/nstemporarydirectory/
        NSURL *groupContainerURL = [[NSFileManager defaultManager] containerURLForSecurityApplicationGroupIdentifier:@"UBF8T346G9.Office"];
        
        sTargetDoc = [NSString stringWithFormat:@"%@_%@", [[NSProcessInfo processInfo] globallyUniqueString], [sOtherDoc lastPathComponent]];
        sTargetDoc = [[groupContainerURL URLByAppendingPathComponent:sTargetDoc] path];
                
        [[word activeDocument] saveAsFileName:sTargetDoc fileFormat:[baseDoc saveFormat] lockComments:NO password:nil addToRecentFiles:NO writePassword:nil readOnlyRecommended:NO embedTruetypeFonts:[baseDoc embedTrueTypeFonts]  saveNativePictureFormat:NO saveFormsData:NO textEncoding:WordMsoEncodingEncodingUTF8 insertLineBreaks:NO allowSubstitutions:NO lineEndingType:WordWdLineEndingTypeLineEndingCrLf HTMLDisplayOnlyOutput:NO maintainCompatibility:YES];
        
        // Close original document
        otherDoc = [word open:nil fileName:sOtherDoc confirmConversions:YES readOnly:NO addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];
        [otherDoc closeSaving:WordSaveOptionsNo savingIn:nil];
    }

    return sTargetDoc;
}

void executeMerge(WordApplication *word, WordDocument *baseDoc, NSString *sBaseDoc, NSString *sTheirDoc, NSString *sMyDoc)
{
    WordDocument *theirDoc, *myDocAfterCompare;
    NSString *sTheirDocAfterCompare, *sMyDocAfterCompare;
    
    theirDoc = baseDoc;

    sTheirDocAfterCompare = executeCompare(word, sBaseDoc, sTheirDoc, @"theirs");

    sMyDocAfterCompare = executeCompare(word, sBaseDoc, sMyDoc, @"mine");
    
    //[theirDoc saveIn:nil as:nil];
    //[myDoc saveIn:nil as:nil];
    
    // No 'activate' method -> comment code
    //[myDoc activate]; //required? just in case
    // We cannot activate the document, so we open it, which should activate it
    myDocAfterCompare = [word open:nil fileName:sMyDocAfterCompare confirmConversions:YES readOnly:NO addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];

    [myDocAfterCompare mergeFileName:sTheirDocAfterCompare];
    
    // Clean-up (this should work, even if the docs are still opened in Word)
    if (![sTheirDoc isEqualToString:sTheirDocAfterCompare])
    {
        [[NSFileManager defaultManager] removeItemAtPath:sTheirDocAfterCompare error:nil];
    }

    if (![sMyDoc isEqualToString:sMyDocAfterCompare])
    {
        [[NSFileManager defaultManager] removeItemAtPath:sMyDocAfterCompare error:nil];
    }
    
    // Built-in three-way merge does not work that nicely
    //[myDoc threeWayMergeLocalDocument:myDoc serverDocument:theirDoc baseDocument:baseDoc favorSource:NO];
}

int main(int argc, const char * argv[]) {
    @autoreleasepool
    {
        if(NSApplicationLoad())
        {
            WordApplication * word;
            NSString *sTheirDoc, *sMyDoc, *sBaseDoc, *sMergedDoc;
            WordDocument *baseDoc;
            
            // Microsoft Office versions for Microsoft Windows OS
            uint vOffice2000 = 9, vOffice2002 = 10,/* vOffice2003 = 11 */
            vOffice2007 = 12, vOffice2010 = 14;
            // WdCompareTarget
            WordWdCompareTarget /* wdCompareTargetSelected = WordWdCompareTargetCompareTargetSelected, */
            /* wdCompareTargetCurrent = WordWdCompareTargetCompareTargetCurrent */
            wdCompareTargetNew = WordWdCompareTargetCompareTargetNew;
            //WordWdMergeTarget wdMergeTargetCurrent = WordWdMergeTargetMergeTargetCurrent;
            
            const char** objArgs = &argv[1];
            int num = argc - 1;
            if (num < 4)
            {
                NSString *basename = @(argv[0]); //[NSString stringWithUTF8String:argv[0]];
                printf("Usage: %s merged.doc theirs.doc mine.doc base.doc\n", [[[basename lastPathComponent] stringByDeletingPathExtension] UTF8String]);
                return 1;
            }
            
            sMergedDoc = @(objArgs[0]);
            sTheirDoc = @(objArgs[1]);
            sMyDoc = @(objArgs[2]);
            sBaseDoc = @(objArgs[3]);
            
            if (![[NSFileManager defaultManager] fileExistsAtPath:sTheirDoc])
            {
                printf("File %s does not exist.  Cannot compare the documents.\n", [sTheirDoc UTF8String]);
                return 1;
            }
            
            if (![[NSFileManager defaultManager] fileExistsAtPath:sMergedDoc])
            {
                printf("File %s does not exist.  Cannot compare the documents.\n", [sMergedDoc UTF8String]);
                return 1;
            }
            
            @try
            {
                word = [SBApplication applicationWithBundleIdentifier:@"com.microsoft.Word"];
            }
             @catch(NSException * e)
            {
                printf("You must have Microsoft Word installed to perform this operation.\n");
                return 1;
            }
            
            // The "visible" property does not exist in this interface
            // [word visible];
            
            // Open the base document
            baseDoc = [word open:nil fileName:sTheirDoc confirmConversions:YES readOnly:NO addToRecentFiles:NO repair:NO showingRepairs:NO passwordDocument:nil passwordTemplate:nil revert:NO writePassword:nil writePasswordTemplate:nil fileConverter:WordWdOpenFormatOpenFormatAuto];
            
            @try
            {
                // Merge into the "My" document
                if ([[word version] intValue] < vOffice2000)
                {
                    // Contrary to the original TortoiseSVN/Git script, we cannot use duck typing -> comment out this line,
                    // as we only support the newer interface below
                    //[baseDoc comparePath:sMergedDoc];
                    printf("Warning: Office versions up to Office 2000 are not officially supported.\n");
                    [baseDoc comparePath:sMergedDoc authorName:@"Comparison" target:wdCompareTargetNew detectFormatChanges:YES ignoreAllComparisonWarnings:YES addToRecentFiles:NO];
                }
                else if ([[word version] intValue] < vOffice2007)
                {
                    [baseDoc comparePath:sMergedDoc authorName:@"Comparison" target:wdCompareTargetNew detectFormatChanges:YES ignoreAllComparisonWarnings:YES addToRecentFiles:NO];
                }
                else if ([[word version] intValue] < vOffice2010)
                {
                    [baseDoc mergeFileName:sMergedDoc];
                }
                else
                {
                    //2010 - handle slightly differently as the basic merge isn't that good
                    //note this is designed specifically for svn 3 way merges, during the commit conflict resolution process
                    executeMerge(word, baseDoc, sBaseDoc, sTheirDoc, sMyDoc);
                }
                
                // Show the merge result
                if ([[word version] intValue] < vOffice2007)
                {
                    [[[word activeDocument] windows][0] setVisible:YES];
                }
                
                // Close the first document
                if (([[word version] intValue] >= vOffice2002) && ([[word version] intValue] < vOffice2010))
                {
                    [baseDoc closeSaving:WordSaveOptionsNo savingIn:nil];
                }
                
                // Show usage hint message
                NSAlert *alert = [[NSAlert alloc] init];
                [alert addButtonWithTitle:@"OK"];
                [alert addButtonWithTitle:@"Cancel"];
                [alert setMessageText:@"OSX Word Merge"];
                [alert setInformativeText:@"You have to accept or reject the changes before\nsaving the document to prevent future problems.\n\nWould you like to see a help page on how to do this?"];
                [alert setAlertStyle:NSInformationalAlertStyle];
                
                if ([alert runModal] == NSAlertFirstButtonReturn) {
                    // OK clicked
                    //NSString *urlString = @"http://office.microsoft.com/en-us/assistance/HP030823691033.aspx"; // URL found in original TSVN script 
                    NSString *urlString = @"https://support.office.com/en-us/article/Review-accept-reject-and-hide-tracked-changes-8af4088d-365f-4461-a75b-35c4fc7dbabd";
                    NSURL *url = [NSURL URLWithString:urlString];
                    if( ![[NSWorkspace sharedWorkspace] openURL:url] )
                    {
                        printf("Failed to open url %s\n",[[url description] UTF8String]);
                    }
                }
            }
            @catch(NSException * e)
            {
                printf("Error running merge (merged: %s, theirs: %s, mine: %s, base: %s\n", [sMergedDoc UTF8String], [sTheirDoc UTF8String], [sMyDoc UTF8String], [sBaseDoc UTF8String]);
                return 1;
            }
        }
    }
    return 0;
}






