var exports = exports || {};
var module = module || { exports };
var exports = exports || {};
var module = module || { exports };
const ss = SpreadsheetApp.getActiveSpreadsheet();
// currently attached spreadsheet https://docs.google.com/spreadsheets/d/1xNCrpcp7Z_gu0YANrNeRsPPIBr2KYdNaGKLU2H9pQW8
function onInstall(e) {
  onOpen(e);
}
// utility function to test out what's stored in the script properties object
// also provides a way to blow away the properties stored

function propertiesTesting() {
  // uncomment below to remove all properties from the properties service
  let scriptProperties = PropertiesService.getScriptProperties();
  let newProps = scriptProperties.deleteAllProperties();
  let obj = newProps.getProperties();
  for (key in obj) {
    Logger.log(obj[key]);
  }

  // temp function to run some evaluations on properties service.
  const properties = scriptProperties.getProperties();
  for (key in properties) {
    Logger.log(properties[key]);
  }
}
function onOpen(e) {
  const menu = SpreadsheetApp.getUi().createMenu('Mail Merge');
  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    menu.addItem('Configure merge options', 'startingPageforStandardMerge');
  } else {
    const properties = PropertiesService.getScriptProperties();
    const merges = properties.getProperty('merges');
    if (merges) {
      menu.addItem('Configure merge options', 'configureMergeOptions');
      menu.addItem('Configure send conditions', 'configureMergeConditions');
      menu.addItem('Configure custom attachment', 'configureCustomAttachment');
      menu.addItem('Preview last merge', 'configureMergePreview');
      menu.addItem('Re-run last merge', 'reRunMerge');
    } else {
      menu.addItem('Configure merge Options', 'startingPageforStandardMerge');
    }
  }
  menu.addSeparator().addItem('Get Help', 'helpPageForMerge');
  menu.addToUi();
}
function getAliases() {
  // provides a list of all aliases associated with the logged in user
  const data = {};
  data.aliases = GmailApp.getAliases();
  var email = Session.getEffectiveUser().getEmail();
  data.aliases.unshift(email);
  return data;
}

function verifyTemplateId({ templateId, elementId }) {
  templateId = templateId || '1mHBCSw3htUKBss-OSvyl3xJfzBYjON-dbewJDayhN-8';
  try {
    const templateFile = DriveApp.getFileById(templateId);
    const mimeType = templateFile.getMimeType();
    const fileName = templateFile.getName();
    // only supporting Google Docs and Google Slides at the moment. Throw an error
    // if they don't match up
    if (
      mimeType === 'application/vnd.google-apps.document' ||
      mimeType === 'application/vnd.google-apps.presentation'
    ) {
      if (templateFile) {
        return [
          null,
          {
            url: templateFile.getUrl(),
            fileType: mimeType,
            elementId,
            fileName,
          },
        ];
      }
    }
    return [
      {
        message:
          'Unsupported file template, please use Google Docs or Google Slides to create your template',
      },
      {
        elementId,
      },
    ];
  } catch (e) {
    return [
      { message: e.message },
      {
        elementId,
      },
    ];
  }
}
function getUserDrafts(refresh) {
  if (refresh === void 0) {
    refresh = false;
  }
  try {
    const drafts = GmailApp.getDrafts();
    let messages = drafts.map(function (draft) {
      // get attachments info separately
      const message = draft.getMessage();
      const attachments = getMessageAttachments(message);
      const attachmentNames = attachments.attachments.map(function (attachment) {
        return attachment.name;
      });
      const rawContent = message.getRawContent();
      const body = message.getBody();
      const inlineImages = attachments.images.length
        ? getInlinedImages(rawContent, attachments.images)
        : [];
      const htmlBody = inlineImages.length
        ? inlineImagesInEmailBody(body, inlineImages)
        : body;
      return {
        to: message.getTo(),
        cc: message.getCc(),
        bcc: message.getBcc(),
        attachments: attachmentNames,
        subject: message.getSubject(),
        body: htmlBody,
        id: draft.getId(),
        originalBody: body,
      };
    });
    if (messages.length === 0) {
      messages = [{ subject: 'No Drafts Found', id: '0' }];
    }
    return { drafts: messages, error: null, refresh };
  } catch (e) {
    return {
      drafts: null,
      error: { message: e.message, stack: e.stack },
      refresh,
    };
  }
}
function getRemainingDailyQuota() {
  const remainingQuota = MailApp.getRemainingDailyQuota();
  return remainingQuota;
}
function storeMailMerge(data) {
  /*
      stores previous mail merge info from user in script properties.
      Allows users to continue where they left off if they need to finish the merge at a later time, or run it again with
      additional data.
      */
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperties(data);
    return scriptProperties.getProperties();
  } catch (e) {
    return 'There was an error writing to the properties service';
  }
}
function getMailMerge() {
  // retrieves stored mail merge data from cache and sends it back to the user
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const data = scriptProperties.getProperties();
    return data;
  } catch (e) {
    return { error: e };
  }
}
// ma
function getHeaders() {
  var dataSheet = ss.getActiveSheet();
  var lastColumn = dataSheet.getLastColumn();
  var headers = dataSheet.getRange(1, 1, 1, lastColumn).getValues();
  return headers;
}

// function initApp gathers necessary data to hydrate the app with cached merge info
function initApp() {
  const headers = getHeaders();
  const { aliases } = getAliases();
  const drafts = getUserDrafts();
  const currentSheet = SpreadsheetApp.getActiveSheet().getSheetId();
  const mergeInfo = getMailMerge();
  return { headers, aliases, drafts, currentSheet, mergeInfo }
}

// kind refers to either preview or merge
function merge(
  kind,
  email,
  sendDrafts,
  recipientsHeader,
  mergeTitle,
  mergeConditions,
  customAttachment,
  currentDate
) {
  // if for some reason mergeTitle is not passed generate a random id for it
  Logger.log('mergeTitle')
  Logger.log(mergeTitle)
  if (!mergeTitle) {
    mergeTitle = uuid().slice(-5);
  }
  try {
    const draft = getDraft(email.id, kind);
    const dataSheet = ss.getActiveSheet();
    const currentSheet = SpreadsheetApp.getActiveSheet().getSheetId();
    const headers = createMergeStatusHeadersIfNotFound([mergeTitle]);
    const dataRange = dataSheet.getDataRange();
    // inline the images
    const matchedImgs = email.body.match(/<img[^>]+>/g);
    // var srcRegex = /src="[^\"]+\"/g;
    // var imgToReplace = [];
    let inlinedImages = {};
    // var emailTemplate = email.body.slice(0);
    if (matchedImgs && matchedImgs.length) {
      if (draft.inlinedImages.length) {
        inlinedImages = draft.inlinedImages.reduce(function (accum, next) {
          const decoded = Utilities.base64Decode(next.content);
          const imageBlob = Utilities.newBlob(decoded);
          imageBlob.setName(next.id);
          imageBlob.setContentType(next.mimeType);
          accum[next.id] = imageBlob;
          return accum;
        }, {});
      }
    }
    const mergeData = {
      subject: email.subject,
      attachments:
        kind === 'preview'
          ? draft.attachments
          : draft.attachments.map(function (attachment) {
            return attachment.blob;
          }),
      to: email.to,
      cc: email.cc,
      bcc: email.bcc,
      htmlBody: draft.body,
      mergeHtml: draft.body.slice(),
      previewHtml: email.body,
      inlineImages: inlinedImages,
    };
    if (email.from === 'donotreply@bc.edu') {
      mergeData.noReply = true;
    } else if (email.from) {
      mergeData.from = email.from;
      mergeData.replyTo = email.from;
    }
    const emailColumn = normalizeHeader(recipientsHeader);
    const mergeStatus = normalizeHeader(`Merge Status - ${mergeTitle}`);
    const objects = getRowsData(dataRange);
    const mergeInfo = {
      type: sendDrafts === 'drafts' ? 'drafts' : 'emails',
      success: [],
      skip: [],
      fail: [],
    };
    const mergePreviewObject = {
      mergePreview: [],
      remainingQuota: getRemainingDailyQuota(),
    };
    const { mergePreview } = mergePreviewObject;

    function skipRow(email, mergeStatus, reasons = [], i) {
      const reason = reasons.join(", ");
      // add the row to the skipped row array
      mergeInfo.skip.push({
        email,
        mergeStatus,
        row: i + 2,
        reason
      });
    }
    // if custom attachment exists need to set up the parent folder to store the merge job for later
    // only do this if mergeType is merge
    if (customAttachment && kind === 'merge') {
      // set up folder structure for custom docs when we process each row later
      const templateFile = DriveApp.getFileById(customAttachment.templateId);
      const folderIterator = templateFile.getParents();
      let parentFolder;
      while (folderIterator.hasNext()) {
        const folder = folderIterator.next();
        parentFolder = folder;
      }
      // create subfolder to keep the merged docs tidy

      const customDocFolder = DriveApp.createFolder(`${customAttachment.originalFileName} merged docs on ${currentDate}`);
      customDocFolder.moveTo(parentFolder);
      // add to the customAttachment object for later use
      customAttachment.templateFile = templateFile;
      customAttachment.parentFolder = customDocFolder;
    }
    // loop through the row data to complete the preview or merge jobs
    const _loop_1 = function (i) {
      const rowData = objects[i];
      let status = '';


      if (typeof rowData[mergeStatus] === 'string') {
        status = rowData[mergeStatus].toLowerCase();
      }
      if (status === 'done' || status === '0') {
        // row was already processed
        skipRow(rowData[emailColumn], rowData[mergeStatus], ['already processed'], i)

      } else {
        try {
          // hang on to some common variables here. track whether all merge conditions are met
          const reasonsSkipped = []; // track the reason for skipping if any
          let isMergeConditionMet = true;


          if (mergeConditions.length) {

            mergeConditions.filter(c => c.currentSheet === currentSheet).forEach(function (_a) {
              // condition contains a column to check and what should be in the cell to match
              // the condition
              const { column, comparison } = _a; // column to be compared against and the comparison chosen
              const { condition } = _a; // text value from the form input on client side

              const normalizedColumn = normalizeHeader(column);
              const cell = rowData[normalizedColumn]; // tired of typing this whole thing out.
              // TODO: need to do better type checking here
              // Cell data can contain number, booleans, dates, strings
              // save variables that attempt to convert the data to a number for the number switch statements
              const cellConvertedToNum = parseFloat(rowData[normalizedColumn]);
              const conditionConvertedToNum = parseFloat(condition);
              // This switch statement is only pushing rows into the skip category, the default is to send all merge data unless
              // it sets the merge condition to false in the code below
              switch (comparison) {
                // cases come from js/index.js where we set the comparison select values
                // #region Text Cases
                case 'TextEquals': {
                  if (rowData[normalizedColumn] !== condition) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} not exact match to ${condition}`)
                  }
                  break;
                }
                case 'TextNotEquals': {
                  if (rowData[normalizedColumn] === condition) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} matched exactly ${condition}`)
                  }
                  break;
                }
                case 'TextEqualsIgnoreCase': {
                  // // first check to see whether contents are of type string
                  // if (typeof rowData[normalizedColumn] !== 'string') {
                  //   isMergeConditionMet = false;
                  //   reasonsSkipped.push(`${cell} did not match ${condition}`)
                  //   break;
                  // }

                  if (rowData[normalizedColumn].toLowerCase() !== condition.toLowerCase()) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} did not match ${condition}`)
                  }
                  break;
                }
                case 'TextNotEqualsIgnoreCase': {
                  // first check to see whether contents are of type string
                  // if (typeof rowData[normalizedColumn] !== 'string') {
                  //   isMergeConditionMet = false;
                  //   reasonsSkipped.push(`${cell} was equal to ${condition}`)
                  //   break;
                  // }

                  if (rowData[normalizedColumn].toLowerCase() === condition.toLowerCase()) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} was equal to ${condition}`)
                  }
                  break;
                }
                case 'TextIsEmpty': {

                  const notEmpty = typeof rowData[normalizedColumn] === 'string' && rowData[normalizedColumn].trim().length !== 0
                  //Logger.log({ notEmpty, header: normalizedColumn, data: rowData[normalizedColumn] })
                  if (notEmpty) {
                    // skip row, merge condition not met for one or more of the conditions set
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${column} was not empty`)
                  }
                  break;
                }
                case 'TextIsNotEmpty': {

                  if (rowData[normalizedColumn] == undefined) {
                    // skip row, merge condition not met for one or more of the conditions set
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} was empty`)
                  } else if (typeof rowData[normalizedColumn] === 'string' && rowData[normalizedColumn].trim().length === 0) {
                    // skip row, merge condition not met for one or more of the conditions set
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${column} was empty`)
                  }
                  break;
                }
                case 'TextContains': {
                  const textFound = rowData[normalizedColumn].includes(condition);
                  if (!textFound) {
                    //skip, shouldn't be in the cell
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${condition} was not found  in ${cell}`)
                  }
                  break;
                }
                case 'TextDoesNotContain': {
                  const textFound = rowData[normalizedColumn].includes(condition);
                  if (textFound) {
                    //skip, shouldn't be in the cell
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${condition} was found  in ${cell}`)
                  }
                  break;
                }
                // #endregion Text Cases
                // #region Num Cases
                case 'numEquals': {


                  if (Number.isNaN(conditionConvertedToNum) || Number.isNaN(cellConvertedToNum)) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} or ${condition} could not be converted  to a number`)
                    break;
                  }
                  if (cellConvertedToNum !== conditionConvertedToNum) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} did not equal ${condition}`)
                  }
                  break;
                }
                case 'numNotEquals': {

                  if (Number.isNaN(conditionConvertedToNum) || Number.isNaN(cellConvertedToNum)) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} or ${condition} could not be converted  to a number`)
                    break;
                  }
                  if (cellConvertedToNum === conditionConvertedToNum) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} did not equal ${condition}`)
                  }
                  break;
                }
                case 'numLessThan': {

                  if (Number.isNaN(conditionConvertedToNum) || Number.isNaN(cellConvertedToNum)) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} or ${condition} could not be converted  to a number`)
                    break;
                  }
                  if (cellConvertedToNum > conditionConvertedToNum) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} was greater than or equal to ${condition}`)
                  }
                  break;
                }
                case 'numGreaterThan': {

                  if (Number.isNaN(conditionConvertedToNum) || Number.isNaN(cellConvertedToNum)) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} or ${condition} could not be converted  to a number`)
                    break;
                  }
                  if (cellConvertedToNum < conditionConvertedToNum) {
                    isMergeConditionMet = false;
                    reasonsSkipped.push(`${cell} was less than or equal to ${condition}`)
                  }
                  break;
                }
                // #endregion Num Cases
                default: {

                  if (rowData[normalizedColumn] != condition) {
                    // skip row, merge condition not met for one or more of the conditions set
                    isMergeConditionMet = false;
                    reasonsSkipped.push('condition not met');
                  }
                }
              }

            });
          }
          if (isMergeConditionMet) {
            const messagePreview = processRow(
              rowData,
              kind,
              mergeData,
              sendDrafts,
              emailColumn,
              headers,
              customAttachment,
              currentDate
            );
            if (messagePreview && kind === 'preview') {
              // just sending back a preview, no need to merge
              mergePreview.push(messagePreview);
            } else {
              // merge was successful set status of merge title column
              dataSheet
                .getRange(
                  i + 2,
                  headers.indexOf(`Merge Status - ${mergeTitle}`) + 1
                )
                .setValue('Done')
                .clearFormat()
                .setBackground('#A3FFDF')
                .setNote(new Date());
              mergeInfo.success.push({
                email: rowData[emailColumn],
                mergeStatus: rowData[mergeStatus],
              });
            }
          } else {
            // mergeCondition not met
            if (kind !== 'preview') {
              const dataRange = dataSheet
                .getRange(
                  i + 2,
                  headers.indexOf(`Merge Status - ${mergeTitle}`) + 1
                );
              dataRange
                .clearFormat()
                .setBackground('#FFE3A3')
              if (reasonsSkipped.length > 0) {
                dataRange.setNote(reasonsSkipped.join(", "));
              }
            }
            skipRow(rowData[emailColumn], rowData[mergeStatus], reasonsSkipped, i);
          }
        } catch (e) {
          dataSheet
            .getRange(
              i + 2,
              headers.indexOf(`Merge Status - ${mergeTitle}`) + 1
            )
            .setValue('Error')
            .setBackground('#ffa7a2')
            .setNote(`${e.message} ${e.stack}`);
          mergeInfo.fail.push({
            email: rowData[emailColumn],
            mergeStatus: e.message,
          });
        }
      }
    };
    for (let i = 0; i < objects.length; ++i) {
      _loop_1(i);
    }
    if (kind === 'preview') return { mergePreviewObject, skipped: mergeInfo.skip };
    return { message: 'merge complete', mergeInfo };
  } catch (e) {
    return { error: e.message, stack: e.stack };
  }
}
function processRow(
  rowData,
  kind,
  mergeData,
  sendDrafts,
  emailColumn,
  headers,
  customAttachment,
  currentDate
) {
  // have to handle preview and merge html templates differently because
  // of the way images are handled in sending emails or creating drafts.
  let emailText = '';
  if (kind === 'preview') {
    emailText = fillInTemplateFromObject(mergeData.previewHtml, rowData);
  } else {
    emailText = fillInTemplateFromObject(mergeData.mergeHtml, rowData);
  }
  const emailSubject = fillInTemplateFromObject(mergeData.subject, rowData);
  let emailTo = fillInTemplateFromObject(mergeData.to, rowData);
  if (emailTo.indexOf(rowData[emailColumn]) === -1) {
    emailTo = emailTo.length
      ? emailTo.concat(', ', rowData[emailColumn])
      : rowData[emailColumn];
  }
  mergeData.htmlBody = emailText;
  if (rowData.cc != undefined) mergeData.cc = rowData.cc;
  if (rowData.bcc != undefined) mergeData.bcc = rowData.bcc;
  if (kind === 'preview') {
    // check for custom attachment and send back a preview of the attachment name for the preview
    const processedPreview = {
      to: emailTo,
      subject: emailSubject,
      body: emailText,
      cc: mergeData.cc,
      bcc: mergeData.bcc,
    };

    if (customAttachment) {

      // just need the name for now
      processedPreview.customAttachment = fillInTemplateFromObject(
        customAttachment.fileName,
        rowData
      );
    }
    return processedPreview;
  }
  if (customAttachment) {
    // request to create custom attachment(s)


    var customPDF = generateCustomPDF(
      customAttachment,
      headers,
      rowData,
    );
  }
  Logger.log('sending drafts or emails')
  Logger.log(sendDrafts)
  if (sendDrafts === 'drafts') {
    GmailApp.createDraft(emailTo, emailSubject, emailText, {
      ...mergeData,
      attachments: customPDF ? [...mergeData.attachments, customPDF] : [...mergeData.attachments],
    });
  } else {
    GmailApp.sendEmail(emailTo, emailSubject, emailText, {
      ...mergeData,
      attachments: customPDF ? [...mergeData.attachments, customPDF] : [...mergeData.attachments],
    });
  }
}
// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance <<Column name>>
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker <<Column name>>
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  if (template === void 0) {
    template = '';
  }
  template = template.replace(/&lt;&lt;/g, '<<');
  template = template.replace(/&gt;&gt;/g, '>>');
  let email = template;
  // Search for all the variables to be replaced, for instance <<Column name>>
  const templateVars = template.match(/<<[^\>]+>>/g);

  if (templateVars != null) {
    // Replace variables from the template with the actual values from the data object.
    // If no value is available, replace with the empty string.
    for (let i = 0; i < templateVars.length; ++i) {
      // normalizeHeader ignores <<>> so we can call it directly here.
      const variableData = data[normalizeHeader(templateVars[i])];
      email = email.replace(templateVars[i], variableData || '');
    }
  }
  return email;
}
// this function will generate custom attachments for each recipient
// needs a doc type (Google Doc or Slide)
// needs the template doc ID
// will generate a PDF
function generateCustomPDF(
  { templateType, templateId, templateName, templateFile, parentFolder: customDocFolder },
  headers,
  mergeData,
) {
  templateType = templateType || 'application/vnd.google-apps.document';
  // Test Doc ID - 1evqsOl84cM_cdCx9ORmrv3O5HvG91w2YHl860xXxw8c
  // Test Presentation ID - 1J506aPFw1clXZ5y18zSUSnKMqFCtXfql_20cpr3w-KM
  templateId = templateId || '1evqsOl84cM_cdCx9ORmrv3O5HvG91w2YHl860xXxw8c';
  headers = headers || ['First Name', 'Last Name', 'Seminar Name'];
  mergeData = mergeData || {
    firstName: 'Kyle',
    lastName: 'Fidalgo',
    seminarName: 'Test Webinar',
  };
  // make sure template name exists and isn't blank
  const newTemplateName =
    templateName && templateName.trim().length > 0
      ? fillInTemplateFromObject(templateName, mergeData)
      : `${mergeData.firstName} ${mergeData.lastName} merge test`;


  // TODO: grab doc title sent from client
  const customPresentationFile = templateFile.makeCopy(
    newTemplateName,
    customDocFolder
  );

  // check for file type, only support GOOGLE_DOCS and GOOGLE_SLIDES for now
  // Docs and slides have different API's so I have to handle each separately
  if (templateType === 'application/vnd.google-apps.document') {
    // TODO : add docs support
    const customDocument = DocumentApp.openById(customPresentationFile.getId());
    // need to replace text in all document sections
    const customBody = customDocument.getBody();
    const customHeader = customDocument.getHeader();
    const customFooter = customDocument.getFooter();
    // need to run this on each section
    // also replaceText doesn't work the same for docs as it does for slides
    // null or undefined values throw an error
    if (customBody) {
      headers.forEach(header => {
        const replacement = mergeData[normalizeHeader(header)];
        customBody.replaceText(`<<${header}>>`, replacement || '');
      });
    }
    if (customHeader) {
      headers.forEach(header => {
        const replacement = mergeData[normalizeHeader(header)];
        customHeader.replaceText(`<<${header}>>`, replacement || '');
      });
    }
    if (customFooter) {
      headers.forEach(header => {
        const replacement = mergeData[normalizeHeader(header)];
        customFooter.replaceText(`<<${header}>>`, replacement || '');
      });
    }
    customDocument.saveAndClose();
  } else if (templateType === 'application/vnd.google-apps.presentation') {
    const customPresentation = SlidesApp.openById(
      customPresentationFile.getId()
    );

    // replace the text with merge data
    // headers has the headers in their original state so we can search for the replacement text
    // normalize headers allows us to grab the data off the merge object by the key name
    headers.forEach(header => {
      customPresentation.replaceAllText(
        `<<${header}>>`,
        mergeData[normalizeHeader(header)]
      );
    });
    customPresentation.saveAndClose();
  } else {
    // mimeType not supported, skip creating PDF
    // TODO: return error here.
    return {
      error: 'Document could not be merged',
    };
  }

  // return as pdf
  return customPresentationFile.getAs('application/pdf');
}
