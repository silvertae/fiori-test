sap.ui.define(
  [
    "sap/ui/core/Fragment",
    "sap/ui/dom/includeScript",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/m/Column",
    "sap/m/ColumnListItem",
    "sap/m/Text"
  ],
  function (Fragment, includeScript, JSONModel, MessageBox, MessageToast, Column, ColumnListItem, Text) {
    "use strict";

    var MAX_FILE_SIZE_BYTES = 10 * 1024 * 1024;
    var MAX_PREVIEW_ROWS = 50;

    function createInitialState() {
      return {
        fileName: "",
        sheetName: "",
        totalRows: 0,
        shownRows: 0,
        columns: [],
        rows: [],
        errors: [],
        errorText: "",
        hasErrors: false,
        canConfirm: false
      };
    }

    return {
      onOpenExcelUploadDialog: async function () {
        var bScriptLoaded = await this._ensureXlsxLibrary();
        if (!bScriptLoaded) {
          return;
        }

        var oDialog = await this._getExcelUploadDialog();
        this._resetUploadState();
        oDialog.open();
      },

      onCloseExcelUploadDialog: function () {
        var oDialog = this._getExistingDialog();
        if (oDialog) {
          oDialog.close();
        }
      },

      onExcelUploadDialogAfterClose: function () {
        this._resetUploadState();
      },

      onExcelFileChange: function (oEvent) {
        this._resetUploadState();

        var oFile = oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        if (!oFile) {
          this._setErrors([this._getText("excelUploadErrorNoFile")]);
          return;
        }

        var sFileName = oFile.name || "";
        this._oUploadModel.setProperty("/fileName", sFileName);

        if (!this._isXlsxFile(sFileName)) {
          this._setErrors([this._getText("excelUploadErrorInvalidExtension")]);
          return;
        }

        if (oFile.size > MAX_FILE_SIZE_BYTES) {
          this._setErrors([this._getText("excelUploadErrorMaxSize")]);
          return;
        }

        this._parseExcelFile(oFile);
      },

      onConfirmExcelUpload: function () {
        var oData = this._oUploadModel.getData();
        if (!oData.canConfirm) {
          return;
        }

        MessageToast.show(
          this._getText("excelUploadSummaryToast", [oData.fileName, String(oData.totalRows), oData.sheetName])
        );

        var oDialog = this._getExistingDialog();
        if (oDialog) {
          oDialog.close();
        }
      },

      _getExcelUploadDialog: function () {
        var oView = this.getView();
        if (!this._pExcelUploadDialog) {
          this._oUploadModel = new JSONModel(createInitialState());
          this._pExcelUploadDialog = Fragment.load({
            id: oView.getId(),
            name: "com.blueward.sample.fioritest.ext.fragment.ExcelUploadDialog",
            controller: this
          }).then(
            function (oDialog) {
              oView.addDependent(oDialog);
              oDialog.setModel(this._oUploadModel, "excelUpload");
              return oDialog;
            }.bind(this)
          );
        }

        return this._pExcelUploadDialog;
      },

      _ensureXlsxLibrary: function () {
        if (window.XLSX) {
          return Promise.resolve(true);
        }

        if (!this._pXlsxScriptLoad) {
          this._pXlsxScriptLoad = includeScript({
            id: "thirdparty-xlsx-script",
            url: sap.ui.require.toUrl("com/blueward/sample/fioritest/thirdparty/xlsx.full.min.js")
          })
            .then(
              function () {
                if (!window.XLSX) {
                  MessageBox.error(this._getText("excelUploadLibraryMissing"));
                  return false;
                }
                return true;
              }.bind(this)
            )
            .catch(
              function () {
                MessageBox.error(this._getText("excelUploadLibraryMissing"));
                return false;
              }.bind(this)
            );
        }

        return this._pXlsxScriptLoad;
      },

      _getExistingDialog: function () {
        return this.byId("excelUploadDialog");
      },

      _parseExcelFile: function (oFile) {
        var oReader = new FileReader();

        oReader.onload = function (oLoadEvent) {
          try {
            var aData = new Uint8Array(oLoadEvent.target.result);
            var oWorkbook = window.XLSX.read(aData, { type: "array" });
            this._applyWorkbookPreview(oWorkbook);
          } catch (oError) {
            this._setErrors([this._getText("excelUploadErrorParseFailed")]);
          }
        }.bind(this);

        oReader.onerror = function () {
          this._setErrors([this._getText("excelUploadErrorReadFailed")]);
        }.bind(this);

        oReader.readAsArrayBuffer(oFile);
      },

      _applyWorkbookPreview: function (oWorkbook) {
        var aSheetNames = oWorkbook.SheetNames || [];
        if (!aSheetNames.length) {
          this._setErrors([this._getText("excelUploadErrorNoSheet")]);
          return;
        }

        var sSheetName = aSheetNames[0];
        var oSheet = oWorkbook.Sheets[sSheetName];
        var aRawRows = window.XLSX.utils.sheet_to_json(oSheet, {
          header: 1,
          defval: "",
          blankrows: false
        });

        if (aRawRows.length < 2) {
          this._setErrors([this._getText("excelUploadErrorNoDataRows")]);
          return;
        }

        var aHeaders = this._normalizeHeaders(aRawRows[0]);
        if (!aHeaders.length) {
          this._setErrors([this._getText("excelUploadErrorNoColumns")]);
          return;
        }

        var aDataRows = aRawRows.slice(1).filter(function (aRow) {
          return aRow.some(function (vCell) {
            return String(vCell).trim() !== "";
          });
        });

        if (!aDataRows.length) {
          this._setErrors([this._getText("excelUploadErrorNoDataRows")]);
          return;
        }

        var aColumns = aHeaders.map(function (sHeader, iIndex) {
          return {
            key: "col" + iIndex,
            label: sHeader
          };
        });

        var aPreviewRows = aDataRows.slice(0, MAX_PREVIEW_ROWS).map(function (aRow) {
          var oRow = {};
          aColumns.forEach(function (oColumn, iIndex) {
            oRow[oColumn.key] = aRow[iIndex] == null ? "" : String(aRow[iIndex]);
          });
          return oRow;
        });

        this._oUploadModel.setData({
          fileName: this._oUploadModel.getProperty("/fileName"),
          sheetName: sSheetName,
          totalRows: aDataRows.length,
          shownRows: aPreviewRows.length,
          columns: aColumns,
          rows: aPreviewRows,
          errors: [],
          errorText: "",
          hasErrors: false,
          canConfirm: true
        });

        this._rebuildPreviewTable(aColumns);
      },

      _normalizeHeaders: function (aHeaders) {
        if (!Array.isArray(aHeaders)) {
          return [];
        }

        return aHeaders.map(function (vHeader, iIndex) {
          var sHeader = String(vHeader == null ? "" : vHeader).trim();
          return sHeader || "Column" + (iIndex + 1);
        });
      },

      _rebuildPreviewTable: function (aColumns) {
        var oTable = this.byId("excelPreviewTable");
        oTable.destroyColumns();
        oTable.unbindItems();

        aColumns.forEach(function (oColumn) {
          oTable.addColumn(
            new Column({
              header: new Text({ text: oColumn.label })
            })
          );
        });

        var aCells = aColumns.map(function (oColumn) {
          return new Text({ text: "{excelUpload>" + oColumn.key + "}" });
        });

        oTable.bindItems({
          path: "excelUpload>/rows",
          template: new ColumnListItem({
            cells: aCells
          })
        });
      },

      _setErrors: function (aErrors) {
        this._oUploadModel.setProperty("/errors", aErrors);
        this._oUploadModel.setProperty("/errorText", aErrors.join("\n"));
        this._oUploadModel.setProperty("/hasErrors", aErrors.length > 0);
        this._oUploadModel.setProperty("/canConfirm", false);
        this._oUploadModel.setProperty("/sheetName", "");
        this._oUploadModel.setProperty("/totalRows", 0);
        this._oUploadModel.setProperty("/shownRows", 0);
        this._oUploadModel.setProperty("/rows", []);
        this._oUploadModel.setProperty("/columns", []);
        this._rebuildPreviewTable([]);
      },

      _resetUploadState: function () {
        if (!this._oUploadModel) {
          return;
        }

        this._oUploadModel.setData(createInitialState(), true);
        this._rebuildPreviewTable([]);

        var oUploader = this.byId("excelFileUploader");
        if (oUploader) {
          oUploader.clear();
        }
      },

      _isXlsxFile: function (sFileName) {
        return /\.xlsx$/i.test(sFileName || "");
      },

      _getText: function (sKey, aArgs) {
        return this.getView()
          .getModel("i18n")
          .getResourceBundle()
          .getText(sKey, aArgs);
      }
    };
  }
);
