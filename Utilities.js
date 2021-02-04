// Compiled using ts2gas 1.6.2 (TypeScript 3.6.4)
var exports = exports || {};
var module = module || { exports: exports };
var exports = exports || {};
var module = module || { exports: exports };
//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////
// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(range) {
    var data = range.getValues();
    var headers = data.shift();
    return getObjects(data, normalizeHeaders(headers));
}
// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            // Logger.log({ cellData, type: typeof cellData, header: keys[j] })
            if (isCellEmpty(cellData)) {
                // might not need this?
                // object[keys[j]] = '';
                // continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    //Logger.log(objects)
    return objects;
}
// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader(headers[i]);
        if (key.length > 0) {
            keys.push(key);
        }
    }
    return keys;
}
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == '<' || letter == '>') {
            continue;
        }
        if (letter == " " && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        }
        else {
            key += letter.toLowerCase();
        }
    }
    return key;
}
// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
    return typeof (cellData) == "string" && cellData == "";
}
// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
    return char >= 'A' && char <= 'Z' || char >= 'a' && char <= 'z' || isDigit(char);
}
// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
    return char >= '0' && char <= '9';
}
function getHeaders() {
    var sheet = ss.getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastColumn).getValues();
    return headers[0];
}
/**
 * @function createMergeStatusHeadersIfNotFound
 * @param {string[]} merges Array of strings containing merge titles that will be appended to the headers row
 * @return {Array}
 */
function createMergeStatusHeadersIfNotFound(merges) {
    var sheet = ss.getActiveSheet();
    var lastColumn = sheet.getLastColumn();
    var headers = getHeaders();
    var mergeStatus = 'Merge Status - ';
    if (!headers.length) {
        merges.forEach(function (merge) { return sheet.getRange(1, lastColumn + merges.length).setValue(mergeStatus + merge); });
        return merges;
    }
    else {
        merges.forEach(function (merge) {
            if (headers.indexOf(mergeStatus + merge) == -1) {
                sheet.getRange(1, headers.length + 1).setValue(mergeStatus + merge);
                headers.push(merge);
            }
        });
    }
    return headers;
}
function getMessageAttachments(message, kind) {
    // given a Gmail message
    // return attachments separated by attachment and inlined images
    // if kind flag is passed as merge also send back blobs
    var imageArray = [];
    var attachmentArray = [];
    message.getAttachments().forEach(function (attachment) {
        var name = attachment.getName();
        var mimeType = attachment.getContentType();
        var content = Utilities.base64Encode(attachment.getBytes());
        if (mimeType.indexOf('image') !== -1) {
            imageArray.push({
                name: name,
                mimeType: mimeType,
                content: content
            });
        }
        else {
            attachmentArray.push({
                name: name,
                mimeType: mimeType,
                content: content,
                blob: kind ? attachment : ''
            });
        }
    });
    return {
        images: imageArray,
        attachments: attachmentArray
    };
}
// var re = /Content-Type:\simage(?:.|\n)*?X-Attachment-Id:\s(.*)/g;
//       let resultsArr = [];
//       const imgIds = [];
//       while ((resultsArr = re.exec(draft.rawContent)) !== null) {
//         imgIds.push(resultsArr[1]);
//       }
//       console.log(draft)
//       console.log(imgIds);
//       const attachmentNames = draft.attachments.filter(attachment => !attachment.mimeType.startsWith('image')).map(attachment => attachment.name)
function getInlinedImages(rawContent, images) {
    //return array of imageIds, image content, mimeType
    // regex to pull only image content types from raw html data
    // captures the id of the image as a capture group
    var re = /Content-Type:\simage(?:.|\n|\r)*?X-Attachment-Id:\s(.*)/g;
    var resultsArr = [];
    var inlineImages = [];
    var _loop_1 = function () {
        resultsArr = re.exec(rawContent);
        if (resultsArr !== null) {
            var regexFilename = new RegExp('Content-Type:\\simage(?:.|\\n|\\r)*?filename="(.*)"', 'g');
            var fileName_1 = regexFilename.exec(resultsArr[0]);
            if (!fileName_1)
                return "continue";
            var image = images.find(function (image) { return image.name === fileName_1[1]; });
            if (image)
                inlineImages.push({ id: resultsArr[1], name: fileName_1, mimeType: image.mimeType, content: image.content });
        }
    };
    // itereate over the results from regex.exec to get each instance of the match
    while (resultsArr !== null) {
        _loop_1();
    }
    return inlineImages;
}
function inlineImagesInEmailBody(body, inlineImages) {
    // an array of matched image tags found
    var emailTemplate = body;
    var matchedImgs = body.match(/<img[^>]+>/g);
    var srcRegex = /src="[^\"]+\"/g;
    if (matchedImgs && matchedImgs.length) {
        matchedImgs.forEach(function (img) {
            var found = img.match(srcRegex)[0];
            var cid = found.substring(found.indexOf(":") + 1, found.length - 1);
            var imageInfo = inlineImages.find(function (image) { return image.id === cid; });
            if (!imageInfo)
                return;
            emailTemplate = emailTemplate.replace(found, "src=\"data:" + imageInfo.mimeType + ";base64," + imageInfo.content + "\"");
        });
    }
    return emailTemplate;
}
function getDraft(draftId, kind) {
    // if kind is merge, grab the blob content of attachments only
    var draft = GmailApp.getDraft(draftId);
    var message = draft.getMessage();
    var _a = getMessageAttachments(message, kind), attachments = _a.attachments, images = _a.images;
    var inlinedImages = images.length ? getInlinedImages(message.getRawContent(), images) : [];
    var messageDetails = {
        attachments: attachments,
        inlinedImages: inlinedImages,
        body: message.getBody()
    };
    return messageDetails;
}
function uuid(a) { return a ? (a ^ Math.random() * 16 >> a / 4).toString(16) : ([1e7] + -1e3 + -4e3 + -8e3 + -1e11).replace(/[018]/g, uuid); }
