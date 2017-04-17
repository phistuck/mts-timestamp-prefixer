// Copyright (c) PhistucK 2017
// This work in in the public domain, or whatever is legal in your country.
// MIT can also be used.

// Prefixes the file names of all of the MTS files in a folder
// with their timestamp.
// Requires avchd2srt-core to be in the PATH environmental variable,
// or in the same folder.

// Run this using cscript.exe.
(function ()
 {
  var fileSystem = new ActiveXObject("Scripting.FileSystemObject");
  var shell = new ActiveXObject("WScript.Shell");

  var extension = ".mts";
  var months =
   [
    "PlaceholderMonth",
    // Taken from avchd2srt-core, but seems like fair use.
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"];

  //                      Monday 23-Jan-2012 13:09:08
  //                             1      2         3     4     5     6
  var subtitlePattern = /^[a-z]+ (\d+)-([a-z]{3})-(\d+) (\d+):(\d+):(\d+) .+$/i;

  function arrayIndexOf(array, item)
  {
   var length = array.length;
   for (var i = 0; i < length; i++)
   {
    if (array[i] === item)
    {
     return i;
    }
   }
   return -1;
  }

  function log(str)
  {
   WScript.Echo(str);
  }

  function endsWith(string, substring)
  {
   return string.lastIndexOf(substring) === string.length - substring.length;
  }

  function getFileList(folder, optionalFilter)
  {
   var files = new Enumerator(folder.files);
   var fileList = [];
   var name;
   for (; !files.atEnd(); files.moveNext()) {
    name = files.item().name;
    if (!optionalFilter || optionalFilter(name))
    {
     fileList.push(name);
    }
   }
   return fileList;
  }

  function isMTS(string)
  {
   return endsWith(string.toLowerCase(), extension);
  }

  function pad(number, minimalCharacterCount, padCharacter)
  {
   minimalCharacterCount = minimalCharacterCount || 2;
   if (padCharacter === null || padCharacter === undefined)
   {
    padCharacter = "0";
   }
   var stringifiedNumber = String(number);
   var stringLength = stringifiedNumber.length;
   if (stringLength > minimalCharacterCount)
   {
    return stringifiedNumber;
   }
   return new Array(minimalCharacterCount - stringLength + 1).join(padCharacter) + stringifiedNumber;
  }

  function processTimestamp(subtitle)
  {
   var matches = subtitle.match(subtitlePattern);
   // log(matches.toString());
   // log(pad(arrayIndexOf(months, matches[2])));
   // log(arrayIndexOf(months, matches[2]));
   // log(matches[2]);
   
   var date =
        pad(matches[3]) + pad(arrayIndexOf(months, matches[2])) + matches[1];
   var time = pad(matches[4]) + pad(matches[5]) + pad(matches[6])
   
   return date + "-" + time;
  }

  function main()
  {
   // var scriptFolder =
   //      fileSystem.getFolder(fileSystem.getParentFolderName(
   //       fileSystem.getFile(WScript.ScriptFullName)));
   var currentFolder = fileSystem.getFolder(shell.currentDirectory);

   var fileList = getFileList(currentFolder, isMTS);
   var fileCount = fileList.length;
   

   var executedCommand;
   var i;

   var shotNamePattern = /^\d+\.mts$/i;
   var processedNameFound = false;
   for (i = 0; i < fileCount; i++)
   {
    if (!fileList[i].match(shotNamePattern))
    {
     processedNameFound = true;
     log("\n\nWarning!");
     log("Some files seem to have already been processed.");
     log("You have about ten seconds to exit before the process begins.");
     log("Example -");
     log(fileList[i]);
     WScript.sleep(10000);
     break;
    }
   }

   var pattern = new RegExp("\\" + extension + "$", "i");
   
   for (i = 0; i < fileCount; i++)
   {
    if (!fileSystem.fileExists(fileList[i].replace(pattern, ".srt")))
    {
     executedCommand = shell.exec("avchd2srt-core " + fileList[i]);
     while (executedCommand.status === 0)
     {
      WScript.sleep(1);
     }
     // log(executedCommand.status);
    }
   }

   var srtFile, line, timestampPrefixedName, currentName;
   
   // log(pattern);
   for (i = 0; i < fileCount; i++)
   {
    currentName = fileList[i];
    // log(fileList[i].replace(pattern, ".srt") + " - " + fileList[i]);
    srtFile = fileSystem.openTextFile(currentName.replace(pattern, ".srt"));
    
    // The SRT file starts with the following pattern.
    // 1
    // 00:00:00,000 --> 00:00:01,040
    // Saturday 14-Jan-2012 18:53:06 (+02:00)

    // Skip the first line - subtitle number.
    srtFile.skipLine();
    // Skip the second line - subtitle timing information.
    srtFile.skipLine();
    // Get the the third line - the timestamp.
    timestampLine = srtFile.readLine();
    // Process the timestamp and append the original name.
    timestampPrefixedName = processTimestamp(timestampLine) + "-" + currentName;
    // Rename the file.
    fileSystem.moveFile(currentName, timestampPrefixedName);
   }
   log("Renamed " + fileCount + " files.");
  }

  main();
 }());