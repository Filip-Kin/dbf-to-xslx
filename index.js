const Parser = require('@filip96/node-dbf');
const createCsvStringifier = require('csv-writer').createObjectCsvStringifier;
const { readdirSync: readdir, createWriteStream, existsSync, mkdirSync } = require('fs');
const { join: pathJoin } = require('path');
const { exec } = require('child_process');
const cliProgress = require('cli-progress')
const config = require('./config.js');

const bar = new cliProgress.SingleBar({
    format: '{bar} | {percentage}% {stage} {file}'
}, cliProgress.Presets.shades_classic);

if (!existsSync('./temp')){
    mkdirSync('./temp');
}

bar.start(1, 0, {
    stage: 'Reading directory',
    file: ''
});

let filesToParse = [];
for (let dir of config.folderPaths) {
    let files = readdir(dir);
    for (let file of files) {
        if (file.toLowerCase().endsWith(config.fileExtension)) {
            let noMatches = true;
            for (let regex of config.ignoreFiles) {
                if (file.match(regex)) {
                    noMatches = false;
                    break;
                }
            }
            if (noMatches) filesToParse.push(pathJoin(dir, file));
        }
    }
}

bar.setTotal(filesToParse.length);

async function processFiles() {
    for (let file of filesToParse) {
        await new Promise((resolve, reject) => {
            let fileName = file.split('\\');
            fileName = fileName[fileName.length - 1].replace(config.fileExtension, '').replace(config.fileExtension.toUpperCase(), '');
            bar.increment(1, {
                stage: 'Reading',
                file: fileName
            });

            let rows, writer, headers = [];
            let parser = new Parser(file);
            let writeStream = createWriteStream(pathJoin('./temp/', fileName + '.csv'));

            parser.on('header', (h) => {
                rows = h.numberOfRecords;
                for (let col of h.fields) {
                    headers.push({
                        id: col.name,
                        title: col.name
                    });
                }
                writer = createCsvStringifier({
                    header: headers
                });
                writeStream.write(writer.getHeaderString());
            });

            parser.on('record', (record) => {
                let rowEmpty = true, oversized = false; row = [{}];
                for (let col of headers) {
                    if (record[col.id] == "") rowEmpty = false;
                    row[0][col.id] = record[col.id];
                }
                bar.update({stage: 'Reading '+record['@sequenceNumber']+'/'+rows});
                if (!rowEmpty || oversized) writeStream.write(writer.stringifyRecords(row));
            });

            parser.on('end', () => {
                writeStream.close();
                bar.update({stage: 'Converting'});
                exec('java -jar csvtoxlsx-jar-with-dependencies.jar "./temp/'+fileName+'.csv" "'+config.exportPath+'/'+fileName+'.xlsx"', (err, out) => {
                    if (err) console.error('\n'+err);
                    if (!out.match('Done!')) console.error('\n'+out);
                    resolve()
                });
                
            });

            parser.parse();
        });
    }
    bar.stop();
}

processFiles();