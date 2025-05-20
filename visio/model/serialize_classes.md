# Visio Document Serialization Class Implementation Table

## Table Description

This table serves as a reference for classes used for serialization/deserialization of Visio documents in SDKJS. It contains the following information:

1. **JavaScript Class (fromXml)** - The name of the JavaScript class that implements the `prototype.fromXml` method. Classes are listed in the order they appear in the `sdkjs-ooxml/visio/model/VisioDocument.js` file. The `AscVisio.` prefix is omitted for readability.Do not add or remove existing names in 1 column

2. **C++ Class (fromPPTY)** - The corresponding C++ class that implements the `fromPPTY` method for deserialization from binary format. An empty cell indicates that there is no corresponding C++ class or that its name is unknown.

3. **Binary Reading Implementation (SerializeReader.js)** - Indicator of the presence or absence of binary format reading method implementations in JavaScript (in the `visio/model/SerializeReader.js` file):
   - "+" means the class has binary reading method implementations (readAttribute/readChild)
   - "-" means the class does not have binary reading method implementations

4. **Binary Writing Implementation (SerializeWriter.js)** - Indicator of the presence or absence of binary format writing method implementations in JavaScript (in the `visio/model/SerializeWriter.js` file):
   - "+" means the class has binary writing method implementations (writeAttribute/writeChild)
   - "-" means the class does not have binary writing method implementations

At the end of the table, C++ classes that have no direct JavaScript counterparts are also listed.

This table is used to track the status of class serialization/deserialization implementations and can serve as a guide when developing new serialization methods.

<!-- ## Passed tasks specification

Look at the any row in the table that have both non-empty first and second column and doesn't have a plus sign in the 3rd column.
Add implementation only of the readChild and readAttribute of the class from the 1st column to the file @[visio/model/SerializeReader.js]. Use the C++ code of the "fromPPTY" function of the class from the 2nd column in the @[VsdxFormat] folder as the basis for the logic.
No need to add init to class members that are declared in @[visio/model/VisioDocument.js] or @[visio/model/ooxmlApi/ooxmlApiIndex.js]
Use Shape_Type.prototype.readAttribute and Shape_Type.prototype.readChild as references for implementation.
Do not check initialized values in the methods (e.g., no need to check "if (!this.authorList) {").
At the end, add a plus sign to this row in the 3rd column. -->

## Current task specification

Look at the any row in the table that have plus sign in the 3rd column.
Add implementation only of the privateWriteAttributes and writeAttribute of the class from the 1st column to the file @[visio/model/SerializeWriter.js]. do not add privateWriteAttributes or writeAttribute with empty implementation. do not check null or undefined values in privateWriteAttributes
Use AscVisio.CWindows.prototype.writeAttribute and AscVisio.CWindows.prototype.writeChild as references for implementation. and AscCommon.CBinaryFileWriter @[sdkjs/common/Drawings/Format/Format.js] as a source of binary methods
Do not check undefined values in the privateWriteAttributes methods (e.g., no need to check "if (this.columnNameID !== undefined) {").
Do not check for null or undefined values before calling WriteRecord2 in writeChildren methods, as WriteRecord2 already has null/undefined checking internally (e.g., no need to check "if (this.documentSheet) { pWriter.WriteRecord2(4, this.documentSheet); }").
Check for null or undefined is needed before calling StartRecord.
At the end, add a plus sign to this row in the 4rd column.


| JavaScript Class (fromXml) | C++ Class (fromPPTY) | Binary Reading Implementation (SerializeReader.js) | Binary Writing Implementation (SerializeWriter.js) |
| --- | --- | --- | --- |
| CVisioDocument | CDocumentFile | + | + |
| CWindows | CWindows | + | + |
| CMasters | CMastersFile | + | + |
| CPageContents | CPageFile | + | + |
| CMasterContents | CMasterFile | + | + |
| CPages | CPagesFile | + | + |
| CComments | CComments | + | + |
| CExtensions | | - | - |
| CDataConnections | CDataConnections | + | + |
| CDataRecordSets | CDataRecordSets | + | + |
| CValidation | CValidationFile | + | + |
| CSolutions | CSolutions | + | + |
| CSolutionXML | | - | - |
| Comments_Type | CCommentsFile | + | + |
| RuleTest_Type | CRuleFormula | + | + |
| RuleFilter_Type | CRuleFormula | + | + |
| RowKeyValue_Type | CRowKeyValue | + | + |
| DataColumn_Type | CDataColumn | + | + |
| RuleInfo_Type | CRuleInfo | + | + |
| IssueTarget_Type | CIssueTarget | + | + |
| Rule_Type | CRule | + | + |
| RuleSetFlags_Type | CRuleSetFlags | + | + |
| AutoLinkComparison_Type | | - | - |
| RefreshConflict_Type | | - | - |
| RowMap_Type | CRowMap | + | + |
| PrimaryKey_Type | CPrimaryKey | + | + |
| DataColumns_Type | CDataColumns | + | + |
| ADOData_Type | | - | - |
| Icon_Type | CIcon | + | + |
| PageSheet_Type | CPageSheet | + | + |
| tp_Type | CText_tp | + | + |
| pp_Type | CText_pp | + | + |
| fld_Type | CText_fld | + | + |
| cp_Type | CText_cp | + | + |
| Rel_Type | | - | - |
| CommentEntry_Type | CCommentEntry | + | + |
| AuthorEntry_Type | CAuthorEntry | + | + |
| RefreshableData_Type | | - | - |
| PublishedPage_Type | | - | - |
| HeaderFooterFont_Type | | - | - |
| FooterRight_Type | | - | - |
| FooterCenter_Type | | - | - |
| FooterLeft_Type | | - | - |
| HeaderRight_Type | | - | - |
| HeaderCenter_Type | | - | - |
| HeaderLeft_Type | | - | - |
| FooterMargin_Type | | - | - |
| HeaderMargin_Type | | - | - |
| CustomToolbarsFile_Type | | - | - |
| CustomMenusFile_Type | | - | - |
| ProtectStyles_Type | | - | - |
| SnapExtensions_Type | | - | - |
| TimePrinted_Type | | - | - |
| TimeEdited_Type | | - | - |
| TimeSaved_Type | | - | - |
| TimeCreated_Type | | - | - |
| CustomProp_Type | | - | - |
| PreviewPicture_Type | | - | - |
| BuildNumberEdited_Type | | - | - |
| BuildNumberCreated_Type | | - | - |
| Template_Type | | - | - |
| AlternateNames_Type | | - | - |
| HyperlinkBase_Type | | - | - |
| Desc_Type | | - | - |
| Keywords_Type | | - | - |
| Category_Type | | - | - |
| Company_Type | | - | - |
| Manager_Type | | - | - |
| Creator_Type | | - | - |
| Subject_Type | | - | - |
| Title_Type | | - | - |
| SectionDef_Type | | - | - |
| FunctionDef_Type | | - | - |
| CellDef_Type | | - | - |
| Issue_Type | CIssue | + | + |
| RuleSet_Type | CRuleSet | + | + |
| ValidationProperties_Type | CValidationProperties | + | + |
| DataRecordSet_Type | CDataRecordSet | + | + |
| DataConnection_Type | CDataConnection | + | + |
| Solution_Type | CSolution | + | + |
| Window_Type | CWindow | + | + |
| Page_Type | CPage | + | + |
| Connect_Type | CConnect | + | + |
| Shape_Type | CShape | + | + |
| MasterShortcut_Type | | - | - |
| Master_Type | CMaster | + | + |
| Text_Type | CText | + | + |
| ForeignData_Type | CForeignData | + | + |
| Data_Type | | - | - |
| RefBy_Type | CRefBy | + | + |
| PublishSettings_Type | | - | - |
| DataTransferInfo_Type | | - | - |
| HeaderFooter_Type | CHeaderFooter | + | + |
| EventItem_Type | CEventItem | + | + |
| DocumentSheet_Type | CDocumentSheet | + | + |
| StyleSheet_Type | CStyleSheet | + | + |
| FaceName_Type | CFaceName | + | + |
| ColorEntry_Type | CColorEntry | + | + |
| DocumentSettings_Type | CDocumentSettings | + | + |
| DocumentProperties_Type | | - | - |
| CellDefBase_Type | | - | - |
| ShapeSheet_Type | | - | - |
| Section_Type | CSection | + | + |
| Trigger_Type | CTrigger | + | + |
| GeometryRow_Type | | - | - |
| IndexedRow_Type | | - | - |
| NamedIndexedRow_Type | | - | - |
| Row_Type | CRow | + | + |
| SolutionXML_Type | | - | - |
| ExtendableCell_Type | | - | - |
| Cell_Type | CCell | + | + |
|  | CText_text | + | + |
|  | CShapes | + | + |
|  | CConnects | - | - |
|  | CEventList | + | + |
|  | CStyleSheets | + | + |
|  | CColors | + | + |
|  | CFaceNames | + | + |
|  | CSolutionFile | - | - |
|  | CSolutionsFile | - | - |
|  | CIssues | - | - |
|  | CRuleSets | - | - |
|  | CCommentList | - | - |
|  | CAuthorList | - | - |
|  | CRuleFormula | - | - |
|  | CConnectionsFile | - | - |
|  | CRecordsetFile | - | - |
|  | CRecordsetsFile | - | - |
