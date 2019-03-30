export default (function() {
  "use strict";
  const winax = require("winax");
  const tmp = require("tmp");
  const win32 = require("./win32");
  const wordConsts = require("./wordConsts");
  

  const wordutils = {
    toArray: function toArray(objlist) {
      const res = [];
      for (let i = 1; i <= objlist.Count; i++) {
        res.push(objlist.Item(i));
      }
      return res;
    }
  };

  const worddoc = function(document) {
    const doc = document;

    const api = {
      close: function close(option) {
        doc.Close(option);
      },
      activate: function activate() {
        doc.Activate;
        win32.setForegroundWindow(doc.Name);
      },
      saveAs: function saveActiveDocumentAsPdf(path, format) {
        doc.SaveAs2(path, format);
      },
      saveAsPdf: function saveActiveDocumentAsPdf(path) {
        doc.SaveAs2(path, wordConsts.WdSaveFormat.wdFormatPDF);
      },
      saveAsRtf: function saveActiveDocumentAsRtf(path) {
        doc.SaveAs2(path, wordConsts.WdSaveFormat.wdFormatRTF);
      },
      setTextfield: function setTextfield(nameOfTextfield, value) {
        console.log("setTextfield: [" + nameOfTextfield + ":" + value + "]");
        wordutils
          .toArray(doc.SelectContentControlsByTag(nameOfTextfield))
          .forEach(textfield => {
            textfield.LockContents = false;
            textfield.Range.Text = value;
            textfield.LockContents = true;
          });
      },
      setCheckbutton: function(nameOfCheckbutton, value) {
        console.log(
          "setCheckbuttonInActiveDocument: [" +
            nameOfCheckbutton +
            ":" +
            value +
            "]"
        );
        wordutils
          .toArray(doc.SelectContentControlsByTag(nameOfCheckbutton))
          .forEach(function(textfield) {
            textfield.LockContents = false;
            textfield.Checked = value;
            textfield.LockContents = true;
          });
      },
      updateTableOfContent: function() {
        wordutils.toArray(doc.TablesOfContents).forEach(toc => {
          toc.Update();
        });
      },
      setNewPagenumberingForLastSection: function() {
        const section = doc.Sections.Last;
        wordutils.toArray(section.Footers).forEach(footer => {
          footer.PageNumbers.RestartNumberingAtSection = true;
          footer.PageNumbers.StartingNumber = 1;
        });
      }
    };

    return api;
  };

  const wordapp = function(isvisible) {

    console.log( 'winax', winax );

    let app;
    let docs;

    const initAppAndDocs = function() {
      if (!app || !app.ProductCode) {
        app = winax.Object("Word.Application");
        app.Visible = isvisible;
        docs = app.Documents;
      }
    };

    const api = {
      setVisible: function(b) {
        app.Visible = b;
      },
      quit: function quit() {
        docs.Count > 0 && docs.Close(false);
        if (app) {
          app.Quit();
          app = null;
        }
      },
      addDocument: function(wordTemplate) {
        initAppAndDocs();
        // Documents.Add Template:= "analyse.dotm", NewTemplate:=False, DocumentType:=wdNewBlankDocument
        return docs.Add(
          wordTemplate,
          false,
          wordConsts.WdNewDocumentType.wdNewBlankDocument
        );
      },
      openDocument: function openDocument(path) {
        initAppAndDocs();
        const doc = docs.Open(path);
        if (isvisible) worddoc(doc).activate();
        return doc;
      },
      getActiveDocument: function getActiveDocument() {
        return app ? app.ActiveDocument : null;
      },
      getDocumentByName: function getDocumentByName(docname) {
        for (let i = 1; i <= docs.Count; i++) {
          if (docs.Item(i).Name.includes(docname)) {
            return docs.Item(i);
          }
        }
      },
      insertImage: (img) => {
        const doc = app.activeDocument;
        const p = doc.inlineshapes.addPicture(img).convertToShape();
        p.WrapFormat.type = 1;
        p.select();
        p.application.selection.insertCaption("Figure", " Das ist ein Test1");

        const pictureName = p.name;
        const captionName = p.application.selection.shapeRange.item(1).name;
        doc.shapes.range(Array(captionName,pictureName)).select();
        const group = doc.application.selection.shapeRange.Group();
//        p.lockAspectRatio = true;
//        p.width = 100;
      }
    };

    const compare = {
      compare: function compare(doc1, doc2, compareDoc) {
        const d1 = api.openDocument(doc1);
        const d2 = api.openDocument(doc2);
        const docCompare = app.CompareDocuments(
          d1,
          d2,
          wordConsts.WdCompareDestination.wdCompareDestinationNew
        );
        d1.close();
        d2.close();
        docCompare.SaveAs(compareDoc);
        docCompare.Close(wordConsts.WdSaveOptions.wdDoNotSaveChanges);
        app.quit();
      }
    };

    const createprodukt = (function() {
      const fcts = {
        pasteFrom: function pasteFrom(docname, selectionJoin) {
          // Documents.Open FileName:="Abschnitt 1.docx", ConfirmConversions:=False, ReadOnly:=False,
          const inDoc = docs.Open(docname, false, true);
          const selectionInDoc = app.Selection;
          selectionInDoc.WholeStory();
          selectionInDoc.Copy();
          inDoc.Close(wordConsts.WdSaveOptions.wdDoNotSaveChanges);
          selectionJoin.Paste();
        },
        createprodukt: function create(
          wordTemplate,
          beitraege,
          anlagen,
          joinDoc
        ) {
          initAppAndDocs();
          const newDoc = docs.Add(
            wordTemplate,
            false,
            wordConsts.WdNewDocumentType.wdNewBlankDocument
          );
          const selection = app.Selection;

          beitraege.forEach(beitrag => {
            selection.EndKey(wordConsts.WdUnits.wdStory); // -> go to end of document
            selection.InsertBreak(wordConsts.WdBreakType.wdPageBreak);
            fcts.pasteFrom(beitrag, selection);
          });

          worddoc(newDoc).updateTableOfContent();

          anlagen.forEach(anlage => {
            selection.EndKey(wordConsts.WdUnits.wdStory); // -> go to end of document
            selection.InsertBreak(wordConsts.WdBreakType.wdSectionBreakNextPage);
            createprodukt.pasteFrom(anlage, selection);

            const footers = newDoc.Sections.Last.Footers;
            wordutils.toArray(footers).forEach(footer => {
              const pageNumbers = footer.PageNumbers;
              pageNumbers.RestartNumberingAtSection = true;
              pageNumbers.StartingNumber = 1;
            });
          });

          //const tempFile = File.createTempFile("AJacobDocumentJoiner-", ".docx");
          newDoc.SaveAs(joinDoc);
          newDoc.Close(wordConsts.WdSaveOptions.wdDoNotSaveChanges);
          app.Quit(wordConsts.WdSaveOptions.wdDoNotSaveChanges);
          //moveTempDocumentToTarget(tempFile, new File(joinDoc));
          return true;
        }
      };
      return {
        createprodukt: fcts.createprodukt
      };
    })();

    return Object.assign({}, api, compare, createprodukt);
  };

  return {
    wordapp,
    worddoc
  };
}());



