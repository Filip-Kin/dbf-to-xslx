exports.folderPaths = [
    "Y:\\Projects\\dbf-to-xslx\\databases" // Converts all files in directory with fileExtension
];
exports.exportPath = "Y:\\Projects\\dbf-to-xslx\\databases",
exports.ignoreFiles =  [
    /Products copy [0-9]/g // Ignore any file that matches regex
];
exports.fileExtension = ".dbf";