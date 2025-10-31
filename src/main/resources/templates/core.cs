            /*
             * メソッド名：z_dcn_export_core
             * 機能概要：BOM帳票出力機能
             * 機能詳細：
             * 備 考 ：アクションまたはDRS/DCNのWF承認で実行される
             * 更新履歴
             * ・SDBMS_2.2.0 Phase2-2：久保野
             * ・SDBMS_2.3.4 改善要望No.96：RGK 錦織
             * ・SDBMS_2.4.X 改善要望No.XX：RGK 萩原
             * ・SDBMS_3.0.2 維持7月 WF見直し RGK錦織
             * ・SDBMS_3.3.1 維持No.24 PMDシート追加　RGK ポーイピュ
             * Copyright(C) 2024 RGK Corporation ALL rights reserved.
             */

            Inn = this.getInnovator();
            ThisCCO = CCO;

            return Main();
        }
        // ================================================
        // 定数
        // ================================================
        private bool _ENABLE_DEBUG_LOG_ = true;      // true=デバッグログを出力する
        private BatchLog.Logger _logger = null;

        // ================================================
        // グローバル変数 / 構造体定義
        // ================================================
        Innovator Inn;
        Aras.Server.Core.CallContext ThisCCO;
        private const string METHOD_NAME = "z_dcn_export_core";
        private string targetinfo = "";
        private BatchLog.Logger batchLogger;

        /// <summary>
        /// アクション名
        /// </summary>
        public readonly struct Actions
        {
            /// <summary>add</summary>
            public const string ADD = "add";
            /// <summary>get</summary>
            public const string GET = "get";
            /// <summary>edit</summary>
            public const string EDIT = "edit";
            /// <summary>delete</summary>
            public const string DELETE = "delete";
            /// <summary>purge</summary>
            public const string PURGE = "purge";
            /// <summary>version</summary>
            public const string VERSION = "version";
        }


        /// <summary>
        /// アイテム管理
        /// </summary>
        private readonly struct Items
        {
            /// <summary>共通プロパティ</summary>
            public readonly struct CmnProps
            {
                public const string ID = "id";
                public const string CONFIG_ID = "config_id";
                public const string SOURCE_ID = "source_id";
                public const string RELATED_ID = "related_id";
                public const string ITEM_NUMBER = "item_number";
                public const string NAME = "name";
                public const string STATE = "state";
                public const string IS_CURRENT = "is_current";
                public const string CLASSIFICATION = "classification";
                public const string TEAM_ID = "team_id";
                public const string MAJOR_REV = "major_rev";
                public const string KEYED_NAME = "keyed_name";

                public const string REV = "z_rev";
                public const string LATEST_EO_NO = "z_latest_eo_no";

            }

            public readonly struct Report
            {
                public const string TYPE_NAME = "z_report_definition";
                public readonly struct Props
                {
                    public const string TEMPLATE_FILE = "z_template_file";

                }

                public readonly struct SheetDef
                {
                    public const string REL_NAME = "z_rel_report_sheet_definition";
                    public readonly struct Props
                    {
                        public const string Z_DETAIL_START_ROW_EXPORT = "z_detail_start_row_export";
                        public const string Z_SHEET_HEIGHT = "z_sheet_height";
                        public const string Z_SHEET_WIDTH = "z_sheet_width";
                        public const string Z_MAX_DETAIL_ROWS = "z_max_detail_rows";
                    }
                }

                public readonly struct ExportDef
                {
                    public const string REL_NAME = "z_rel_report_export_definition";
                    public readonly struct Props
                    {

                    }
                }

            }
        }

        private readonly struct ExportBlockDef
        {
            public const string BASIC_INFO = "1";

            public const string PART_INFO_LABEL1 = "11";
            public const string PART_INFO_VALUE1 = "12";
            public const string PART_INFO_LABEL2 = "21";
            public const string PART_INFO_VALUE2 = "22";
            public const string PART_INFO_LABEL3 = "31";
            public const string PART_INFO_VALUE3 = "32";
            public const string PART_INFO_LABEL4 = "41";
            public const string PART_INFO_VALUE4 = "42";
            public const string PART_INFO_LABEL5 = "51";
            public const string PART_INFO_VALUE5 = "52";
            public const string PART_INFO_LABEL6 = "61";
            public const string PART_INFO_VALUE6 = "62";
            public const string PART_INFO_LABEL7 = "71";
            public const string PART_INFO_VALUE7 = "72";
            public const string PART_INFO_LABEL8 = "81";
            public const string PART_INFO_VALUE8 = "82";

            public const string BOM_INFO1_LABEL = "21";
            public const string BOM_INFO1_VALUE = "22";
            public const string BOM_INFO2_LABEL = "31";
            public const string BOM_INFO2_VALUE = "32";
            public const string BOM_INFO3_LABEL = "41";
            public const string BOM_INFO3_VALUE = "42";

            public const string HISTORY_LABEL = "21";
            public const string HISTORY_VALUE = "22";
        }

        /// <summary>
        /// シート定義インデックス
        /// </summary>
        private readonly struct ExportSheetIndexDef
        {
            public const string DCN = "1";
            public const string PART = "2";
            public const string BOM = "3";
        }

        private readonly struct ExportSheetName
        {
            public const string DCN = "Cover";
            public const string PART = "Basic Info";
            public const string BOM = "BOM";
        }

        /// <summary>
        /// 改行コード
        /// </summary>
        private readonly struct NewLineChars
        {
            /// <summary>
            /// Carriage return
            /// </summary>
            public const string CR = "\r";
            /// <summary>
            /// Line feed
            /// </summary>
            public const string LF = "\n";
            /// <summary>
            /// Carrige Return / Line Feed
            /// </summary>
            public const string CRLF = CR + LF;
        }

        private static Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");

        private Dictionary<string, List<string>> Activities = new Dictionary<string, List<string>>()
        {
            { "Create",                new List<string>() { "作成者",             "Prepared by"           } },
            { "Review1",               new List<string>() { "点検者1",            "Reviewer 1"            } },
            { "Review2",               new List<string>() { "点検者2",            "Reviewer 2"            } },
            { "Review3",               new List<string>() { "点検者3",            "Reviewer 3"            } },
            { "Review4",               new List<string>() { "点検者4",            "Reviewer 4"            } },
            { "Review5",               new List<string>() { "点検者5",            "Reviewer 5"            } },
            { "Review6",               new List<string>() { "点検者6",            "Reviewer 6"            } },
            { "Review7",               new List<string>() { "点検者7",            "Reviewer 7"            } },
            { "SW_PP Inspect",         new List<string>() { "製造生管点検者",     "SW PP Inspector"       } },
            { "SW_PE Inspect",         new List<string>() { "製造生技点検者",     "SW PE Inspector"       } },
            { "Eng.QA Inspect",        new List<string>() { "SD 品証点検者",      "SD QA Inspector"       } },
            { "SW_QA Inspect",         new List<string>() { "SW 品証点検者",      "SW QA Inspector"       } },
            { "SCM Inspect",           new List<string>() { "SCM点検者",          "SCM Inspector"         } },
            // SDBMS_3.0.2 維持7月 WF見直し RGK錦織 START
            //{ "Integration T Inspect", new List<string>() { "統合T点検者",        "Integration T Inspector" } },
            { "Weight T Inspect", new List<string>() { "重量T点検者",        "Weight T Inspector" } },
            // SDBMS_3.0.2 維持7月 WF見直し RGK錦織 END
            { "Manager Review",        new List<string>() { "マネージャー点検者", "Manager Reviewer"      } },
            { "CE Office Approval",    new List<string>() { "CE室承認者",         "CE Office Approver"    } }
        };

        private readonly List<string> EnJaProps = new List<string>()
        {
            "z_assigned_creator",
            "z_designated_user",
            "z_reason_for_change",
            "z_pmd",

            "z_st_inorganic_desc_jp",
            "z_st_organic_desc_jp",
            "z_design_responsibility",
            "z_manufacturing_responsibility",
            "z_mng_interchange_code_label",
            "z_mng_spare_part_class_label",
            "description"
        };


        private Item ReportDefinition = null;
        private Item SheetDefinitions = null;
        private Item ExportDefinition = null;
        private Item SheetDefinition = null;

        private Item BatchSetting = null;
        private Item BatchSettingProcess = null;
        private string TemplateFilePath = "";
        private string TempFolderPath = "";

        private ExcelWorksheets Worksheets;
        private ExcelWorksheet ExportSheet;
        private Point CurrentPosition;
        private int TotalPageCount = 0;
        private int CurrentPageCount = 0;
        private int pageCount = 0;
        private int[] PageCountList = null;
        private int currentPageCount = 0;
        private int TemplateSheetIndex = 1;

        private List<ExcelPicture> excelPictures = new List<ExcelPicture>();     //テンプレートファイルからコピーする画像

        string errMsg = "";

        private Item DCN = null;
        private Item newPart = null;
        private Item oldPart = null;
        //private Item Part = null;
        //private List<Item> PartList = new List<Item>();
        private FileInfo TemplateFile = null;
        private FileInfo ExportFile = null;
        private ExcelPackage ExportExcel = null;
        
        /// <summary>
        /// PDFファイル情報クラス
        /// </summary>
        private class PdfFileInfoClass
        {
            public string FileFolder { get; set; }
            public string FilePath { get; set; }
            public string FileName { get; set; }
        }

        private class ParentItemClass
        {
            public string Id { get; set; }
            public string ItemTypeName { get; set; }

        }

        /// <summary>
        /// メイン処理
        /// </summary>
        private Item Main()
        {
            targetinfo = string.Format("[{0} ID：{1}]", this.getType(), this.getID());
            try
            {
                // ログ設定を取得
                batchLogger = new BatchLog.Logger(Inn, ThisCCO, "z_dcn_log_definition");

                // 開始ログ出力
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "DCN Export開始");

                // 定義アイテムを取得
                InitDefinitions();

                // WorkFlow経由かアクションか判定
                bool isWorkflow = false;
                if (this.getProperty("is_workflow", "0").Equals("1")) isWorkflow = true;
                //debug
                //isWorkflow = true;

                string thisID = this.getID();
                // debug
                //string thisID = "32C6B849322243FF92915692BAB8AC5D";
                //string dcnNo = this.getProperty("z_dcn_no", "");
                // debug
                DCN = Inn.getItemById("z_dcn", thisID);
                string dcnNo = DCN.getProperty("z_dcn_no", "");

                Item affItem = Inn.newItem("z_rel_dcn_affected_item", Actions.GET);
                affItem.setProperty("source_id", thisID);
                affItem = affItem.apply();


                // 対象アイテムからPartを取得
                int affCount = affItem.getItemCount();
                string newPartID = "";
                string oldPartID = "";
                for (int i = 0; i < affCount; i++)
                {
                    // 新アイテムIDの取得
                    string newNum = affItem.getItemByIndex(i).getProperty("z_new_num", "");
                    if (!string.IsNullOrEmpty(newNum))
                    {
                        // Partアイテムの取得
                        Item getItem = Inn.newItem("Part", Actions.GET);
                        getItem.setID(newNum);
                        getItem = getItem.apply();

                        // 取得できた場合は保存
                        if (getItem.getItemCount() == 1)
                        {
                            newPart = getItem;
                            newPartID = newNum;
                        }
                    }
                    // 旧アイテムIDの取得
                    string oldNum = affItem.getItemByIndex(i).getProperty("z_old_num", "");
                    if (!string.IsNullOrEmpty(oldNum))
                    {
                        // Partアイテムの取得
                        // Partアイテムの取得
                        Item getItem = Inn.newItem("Part", Actions.GET);
                        getItem.setID(newNum);
                        getItem = getItem.apply();

                        // 取得できた場合は保存
                        if (getItem.getItemCount() == 1)
                        {
                            oldPart = getItem;
                            oldPartID = oldNum;
                        }
                    }
                }

                // Part Number取得
                // string partNo = newPart.getProperty("item_number", "");
                string partNo = DCN.getProperty("z_part_number", "");
                // Part Type取得
                string classification = newPart.getProperty("classification", "");

                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "作業フォルダ・ファイルを作成");

                // 作業フォルダ作成
                string guid = Guid.NewGuid().ToString("N").ToUpperInvariant();
                TempFolderPath += "\\" + dcnNo + "_" + partNo + guid;
                Directory.CreateDirectory(TempFolderPath);

                // 作業ファイル
                string exportFileName = dcnNo + "_" + partNo + ".xlsx";

                string templateFileName = string.Format("DCN_Template_{0}_{1}.xlsx", Inn.getUserID(), DateTime.Now.ToString("yyyyMMddHHmmssfff"));

                // テンプレートファイルを作業用フォルダにダウンロード
                TemplateFilePath = TempFolderPath + "\\" + templateFileName;
                ReportDefinition.fetchFileProperty(Items.Report.Props.TEMPLATE_FILE, TemplateFilePath, FetchFileMode.Normal);

                // テンプレートファイルを開く
                TemplateFile = new FileInfo(TemplateFilePath);
                if (!TemplateFile.Exists)
                {
                    errMsg = "帳票テンプレートの取得に失敗しました。帳票作成処理を中断します。";
                    throw new Exception(errMsg);
                }

                // 出力するエクセルを作成
                string exportFilePath = TempFolderPath + "\\" + exportFileName;
                ExportFile = new FileInfo(exportFilePath);

                if (ExportFile.Exists)
                {
                    errMsg = "他のプロセスと作業ファイルが競合しています。帳票作成処理を中断します。";
                    throw new Exception(errMsg);
                }

                // 出力するエクセルを開く
                ExportExcel = new OfficeOpenXml.ExcelPackage(ExportFile, TemplateFile);
                Worksheets = ExportExcel.Workbook.Worksheets;

                int sheetIndex = 0;

                // ============================================================= DCN_表紙の出力 =============================================================

                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "DCN_表紙シート出力");

                // シート定義の取得
                SheetDefinition = GetSheetDefinitionsAt(ExportSheetIndexDef.DCN);
                sheetIndex = int.Parse(ExportSheetIndexDef.DCN);

                // 作業用ワークシートを追加
                ExportSheet = Worksheets.Add(ExportSheetName.DCN, Worksheets[SheetDefinition.getProperty("z_sheet_name", "")]);

                // ページカウントを設定
                CurrentPageCount = 0;

                // 操作中セルを設定
                SetCurrentPosition(1, 1);

                // DCN情報を出力
                OutputDCNInfo();

                // 承認履歴を出力
                OutputHistory(thisID);

                // 書式・印刷設定
                SetPrintSetting(int.Parse(ExportSheetIndexDef.DCN));

                // ロゴの出力
                OutputLogo();

                // ページ数の保持
                PageCountList[sheetIndex - 1] = CurrentPageCount;

                // ============================================================= Part基本情報の出力 =============================================================

                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "基本情報シート出力");

                // シート定義の取得
                SheetDefinition = GetSheetDefinitionsAt(ExportSheetIndexDef.PART);
                sheetIndex = int.Parse(ExportSheetIndexDef.PART);

                // 作業用ワークシートを追加
                ExportSheet = Worksheets.Add(ExportSheetName.PART, Worksheets[SheetDefinition.getProperty("z_sheet_name", "")]);

                // ページカウントを設定
                CurrentPageCount = 0;

                // 操作中セルを設定
                SetCurrentPosition(1, int.Parse(SheetDefinition.getProperty("z_detail_start_row_export", "1")));

                // 引数を作成
                string[] procArgs = new string[] { newPartID, oldPartID };

                // Part基本情報の差分出力
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Par基本情報出力");
                Item outputDataList = CallProcedure("z_get_is_was_basic_for_report", procArgs);
                OutputPartDiffInfo(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL1, ExportBlockDef.PART_INFO_VALUE1, outputDataList);

                // Assemblyプロパティの差分出力
                if (classification.Equals("Assembly"))
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Assembly情報出力");
                    OutputPartDiffInfo(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL2, ExportBlockDef.PART_INFO_VALUE2, outputDataList);
                }

                // Componentプロパティの差分出力
                if (classification.Equals("Component"))
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Component情報出力");
                    OutputPartDiffInfoVariable(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL3, ExportBlockDef.PART_INFO_VALUE3, outputDataList);
                }

                // SCDプロパティの差分出力
                if (classification.Equals("SCD"))
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "SCD情報出力");
                    OutputPartDiffInfoVariable(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL4, ExportBlockDef.PART_INFO_VALUE4, outputDataList);
                }

                // Managementタブの差分出力
                if (!classification.Equals("Other") && !classification.Equals("Plan DWG Part"))// SDBMS_2.3.4 改善要望No.96 Layout→Plan DWG Part変更 RGK錦織
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Management情報出力");
                    OutputPartDiffInfo(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL5, ExportBlockDef.PART_INFO_VALUE5, outputDataList);
                }

                // Noteタブの差分出力
                Item outputRelDataList = CallProcedure("z_get_is_was_notes_for_report", procArgs);
                if (0 < outputRelDataList.getItemCount())
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Noteタブ情報出力");
                    //OutputPartDiffInfo(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL6, ExportBlockDef.PART_INFO_VALUE6, outputRelDataList);
                    OutputNoteDiff(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL6, ExportBlockDef.PART_INFO_VALUE6, outputRelDataList, false);
                }

                // SPECタブの差分出力
                outputRelDataList = CallProcedure("z_get_is_was_spec_for_report", procArgs);
                if (0 < outputRelDataList.getItemCount())
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "SPECタブ情報出力");
                    OutputSpecDiff(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL7, ExportBlockDef.PART_INFO_VALUE7, outputRelDataList, false);
                }

                // Long Descriptionの差分出力
                if (0 < outputDataList.getItemCount())
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Long Description出力");
                    OutputPartDiffInfoVariable(ExportSheetIndexDef.PART, ExportBlockDef.PART_INFO_LABEL8, ExportBlockDef.PART_INFO_VALUE8, outputDataList);
                }

                // 書式・印刷設定
                SetPrintSetting(sheetIndex);

                // ロゴの出力
                OutputLogo();

                // ページ数の保持
                PageCountList[sheetIndex - 1] = CurrentPageCount;


                // ============================================================= BOMの出力 =============================================================


                // 構成部品シートの作成(Assembly,Plan DWG Partのみ)
                if (classification.Equals("Assembly") || classification.Equals("Plan DWG Part"))// SDBMS_2.3.4 改善要望No.96 Layout→Plan DWG Part変更 RGK錦織
                {
                    // 出力データ取得
                    Item bomDataList = CallProcedure("z_get_is_was_bom_for_report", procArgs);

                    int dataCount = bomDataList.getItemCount();
                    // 構成部品データがある場合のみ出力
                    if (0 < dataCount)
                    {
                        // BOMシートの定義取得
                        SheetDefinition = GetSheetDefinitionsAt(ExportSheetIndexDef.BOM);
                        sheetIndex = int.Parse(ExportSheetIndexDef.BOM);


                        // 出力値が存在する場合のみ出力
                        if (!IsSkipOutput(bomDataList, ExportSheetIndexDef.BOM, ExportBlockDef.BOM_INFO1_VALUE, ExportBlockDef.BOM_INFO2_VALUE, ExportBlockDef.BOM_INFO3_VALUE))
                        {
                            // 作業用ワークシートを追加
                            ExportSheet = Worksheets.Add(ExportSheetName.BOM, Worksheets[SheetDefinition.getProperty("z_sheet_name", "")]);

                            // ページカウントを設定
                            CurrentPageCount = 0;

                            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "BOM情報出力");

                            // 操作中セルを設定
                            SetCurrentPosition(1, int.Parse(SheetDefinition.getProperty("z_detail_start_row_export", "1")));

                            // BOM1段目を出力
                            OutputBOM(2, ExportSheetIndexDef.BOM, ExportBlockDef.BOM_INFO1_LABEL, ExportBlockDef.BOM_INFO1_VALUE, bomDataList, true, false);

                            // BOM2段目を出力
                            OutputBOM(2, ExportSheetIndexDef.BOM, ExportBlockDef.BOM_INFO2_LABEL, ExportBlockDef.BOM_INFO2_VALUE, bomDataList, true, false);

                            // BOM3段目を出力
                            OutputBOM(2, ExportSheetIndexDef.BOM, ExportBlockDef.BOM_INFO3_LABEL, ExportBlockDef.BOM_INFO3_VALUE, bomDataList, true, true);

                            // 書式・印刷設定
                            SetPrintSetting(sheetIndex);

                            // ロゴの出力
                            OutputLogo();

                            // ページ数の保持
                            PageCountList[sheetIndex - 1] = CurrentPageCount;
                        }
                    }
                }

                // ページ数を出力
                CurrentPageCount = 0;
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "総ページ:" + TotalPageCount.ToString());

                // ページ数を出力
                // SDBMS_3.3.1 維持10月 PMD シート追加 RGK ポーイピュ START
                // DCN「さらに表示」ボタンの「DCN Export」を押下した時だけページ数を出力
                string pmdFlag = DCN.getProperty("z_pmd_sheet", "");
                if(!isWorkflow)
                {
                    ExportSheet = Worksheets["Cover"];
                    OutputPageNumber(37, 1, dcnNo);
                    ExportSheet = Worksheets["Basic Info"];
                    if (ExportSheet != null) OutputPageNumber(36, 2, dcnNo);
                    ExportSheet = Worksheets["BOM"];
                    if (ExportSheet != null) OutputPageNumber(72, 3, dcnNo);
                    //OutputPageNumber();
                }
                // SDBMS_3.3.1 SDBMS_3.1.2 維持10月 PMD シート追加 RGK ポーイピュ END

                // ロゴの出力
                //OutputLogo();

                // 不要なシートを削除
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "テンプレートシートを削除");
                ExportExcel.Compatibility.IsWorksheets1Based = false;
                Worksheets.Delete(0);
                Worksheets.Delete(0);
                Worksheets.Delete(0);
                ExportExcel.Compatibility.IsWorksheets1Based = true;

                // テンプレートファイルを削除
                if (TemplateFile != null) TemplateFile.Delete();

                // 帳票ファイルを出力
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "ファイル保存");
                ExportExcel.Save();

                // ================================================
                // WF承認経由の場合
                // ================================================
                if (isWorkflow)
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "WorkFlow経由");

                    string tempFolder = Path.Combine(Path.GetTempPath(), "z_dcn_export_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(tempFolder);
                    TempFolderPath = tempFolder;

                    string dcnPdfFolder = Path.Combine(tempFolder, "DCN_PDF");
                    Directory.CreateDirectory(dcnPdfFolder);

                    string dcnExcelCopyPath = Path.Combine(dcnPdfFolder, ExportFile.Name);
                    File.Copy(ExportFile.FullName, dcnExcelCopyPath, true);
                    LogDebug("dcnExcelCopyPath: " + dcnExcelCopyPath);

                    try
                    {
                        // PDF変換
                        string dcnParam = "<dir_path>" + dcnPdfFolder + "</dir_path>"
                                        + "<is_delete_original_file>false</is_delete_original_file>";
                        LogDebug("dcnParam: " + dcnParam);

                        Item dcnPdfRes = Inn.applyMethod("z_doc_to_pdf_verification", dcnParam);
                        LogDebug("PDF変換結果: " + dcnPdfRes.getProperty("result", ""));
                        LogDebug("生成PDF一覧: " + dcnPdfRes.getProperty("result_json", ""));

                        if (!dcnPdfRes.getProperty("result", "false").Equals("true"))
                            throw new Exception("DCN ExcelからPDF変換に失敗しました。");

                        string[] dcnPdfPaths = dcnPdfRes.getProperty("result_json", "").Split(',');
                        string oriPDFfile = dcnPdfPaths[0].Trim('[', ']', '"').Trim();

                        if (!File.Exists(oriPDFfile))
                            throw new Exception("生成されたPDFファイルが存在しません: " + oriPDFfile);

                        string finalPdfPath = oriPDFfile;

                        // PMD Sheet結合
                        if (pmdFlag.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                        {
                            Item parentItem = Inn.getItemById("z_dcn", thisID);
                            if (parentItem.getItemCount() != 1)
                                throw new Exception("DCNアイテムが単一で取得できません。ID=" + thisID);

                            Item vaultList = GetTargetItemVaultList(parentItem);
                            string pmdPdfPath = null;
                            for (int i = 0; i < vaultList.getItemCount(); i++)
                            {
                                Item relItem = vaultList.getItemByIndex(i);
                                Item fileItem = relItem.getPropertyItem("related_id");
                                if (fileItem != null && fileItem.getProperty("filename", "").Equals("PMD Sheet.pdf", StringComparison.OrdinalIgnoreCase))
                                {
                                    PdfFileInfoClass pdfInfo = DownloadPdfFile(TempFolderPath, fileItem);
                                    pmdPdfPath = pdfInfo.FilePath;
                                    break;
                                }
                            }

                            if (string.IsNullOrEmpty(pmdPdfPath))
                                throw new Exception("PMD Sheet.pdfがAttachmentsタブに登録されていません。");

                            string combinedPaths = oriPDFfile + "," + pmdPdfPath;
                            string combineParam = "<id>" + DCN.getID() + "</id>"
                                                + "<itemTypeName>z_dcn</itemTypeName>"
                                                + "<oriFile>" + combinedPaths + "</oriFile>";

                            Item combinePDF = Inn.applyMethod("z_dcn_pmd_pdf_combine", combineParam);
                            if (!combinePDF.getProperty("result", "false").Equals("true"))
                                throw new Exception("DCN + PMD PDFマージに失敗しました。");

                            finalPdfPath = combinePDF.getProperty("result_json", "").Trim('[', ']', '"').Trim();
                            if (!File.Exists(finalPdfPath))
                                throw new Exception("結合後のPDFファイルが存在しません: " + finalPdfPath);

                            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "DCN + PMD PDFマージ完了");
                        }

                        try
                        {
                            StampPdf(finalPdfPath, DCN.getProperty("keyed_name"));
                        }
                        catch (Exception ex)
                        {
                            batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, "スタンプ処理失敗: " + ex.Message);
                        }

                        string uploadDir = Path.Combine(Path.GetTempPath(), "z_dcn_upload");
                        Directory.CreateDirectory(uploadDir);

                        string uploadFilePath = Path.Combine(uploadDir, Path.GetFileName(finalPdfPath));
                        File.Copy(finalPdfPath, uploadFilePath, true);

                        Item uploadItem = Inn.newItem("z_upload_download_file", Actions.ADD);
                        uploadItem.setProperty("z_file_name", pmdFlag.Equals("Yes", StringComparison.OrdinalIgnoreCase) ? "DCN_PMD_Sheet_Merged.pdf" : ExportFile.Name);
                        uploadItem.setFileProperty("z_file", uploadFilePath);
                        uploadItem = uploadItem.apply();
                        if (uploadItem.isError())
                            throw new Exception(uploadItem.getErrorString());

                        DCN.setAction(Actions.EDIT);
                        DCN.setAttribute("serverEvents", "0");
                        DCN.setProperty("z_dcn_file", uploadItem.getProperty("z_file", ""));
                        DCN = DCN.apply();
                        if (DCN.isError())
                            throw new Exception(DCN.getErrorString());

                        return this;
                    }
                    finally
                    {
                        DeleteWorkDirectory(); // 例外時も削除
                        batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "PDF処理完了");
                    }
                }
                // if (isWorkflow)
                // {
                //     // SDBMS_3.3.1 維持10月 PMD シート追加 RGK ポーイピュ START
                //     // WorkFlow経由でDCNのファイルを出力
                //     batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO,"WorkFlow経由");

                //     // PDF変換やファイル結合で使用する一時フォルダ作成
                //     string tempFolder = Path.Combine(Path.GetTempPath(), "z_dcn_export_" + Guid.NewGuid().ToString("N"));
                //     Directory.CreateDirectory(tempFolder);
                //     TempFolderPath = tempFolder;

                //     // DCN_PDF用サブフォルダ作成
                //     string dcnPdfFolder = Path.Combine(tempFolder, "DCN_PDF");
                //     Directory.CreateDirectory(dcnPdfFolder);

                //     // ExcelファイルをPDF変換用フォルダへコピー
                //     string dcnExcelCopyPath = Path.Combine(dcnPdfFolder, ExportFile.Name);
                //     File.Copy(ExportFile.FullName, dcnExcelCopyPath, true);

                //     // Excel → PDF 変換処理
                //     LogDebug(" Excel → PDF 変換処理 START ---------------");
                //     string dcnParam = "<dir_path>" + dcnPdfFolder + "</dir_path>";
                //     Item dcnPdfRes = Inn.applyMethod("z_doc_to_pdf_verification", dcnParam);
                //     Item dcnPdfRes = Inn.applyMethod("z_doc_to_pdf_verification", dcnParam);
                //     if (!dcnPdfRes.getProperty("result", "false").Equals("true"))
                //         throw new Exception("DCN ExcelからPDF変換に失敗しました。");

                //     string[] dcnPdfPaths = dcnPdfRes.getProperty("result_json", "").Split(',');
                //     string oriPDFfile = dcnPdfPaths[0].Trim('[', ']', '"').Trim();

                //     string finalPdfPath = oriPDFfile;

                //     // PMD SheetのドロップダウンがYesの場合、PMD Sheetも含めて結合
                //     if (pmdFlag.Equals("Yes", StringComparison.OrdinalIgnoreCase))
                //     {
                //         Item parentItem = Inn.getItemById("z_dcn", thisID);
                //         if (parentItem.getItemCount() != 1)
                //         {
                //            throw new Exception("DCNアイテムが単一で取得できません。ID=" + thisID);
                //         }

                //         // AttachmentsからPDF取得
                //         Item vaultList = GetTargetItemVaultList(parentItem);
                //         string pmdPdfPath = null;

                //         for (int i = 0; i < vaultList.getItemCount(); i++)
                //         {
                //             Item relItem = vaultList.getItemByIndex(i);
                //             Item fileItem = relItem.getPropertyItem("related_id");
                //             if (fileItem != null)
                //             {
                //                 string fileName = fileItem.getProperty("filename", "");
                //                 if (fileName.Equals("PMD Sheet.pdf", StringComparison.OrdinalIgnoreCase))
                //                 {
                //                     PdfFileInfoClass pdfInfo = DownloadPdfFile(TempFolderPath, fileItem);
                //                     pmdPdfPath = pdfInfo.FilePath;
                //                     break;
                //                 }
                //             }
                //         }
                //         if (string.IsNullOrEmpty(pmdPdfPath))
                //         {
                //             throw new Exception("PMD Sheet.pdfがAttachmentsタブに登録されていません。");
                //         }

                //         // DCN + PMD PDF結合
                //         string combinedPaths = oriPDFfile;
                //         if (!string.IsNullOrEmpty(pmdPdfPath))
                //             combinedPaths += "," + pmdPdfPath;

                //         string combineParam = "<id>" + DCN.getID() + "</id>"
                //                                 + "<itemTypeName>z_dcn</itemTypeName>"
                //                                 + "<oriFile>" + combinedPaths + "</oriFile>";

                //         // DCN + PMD PDF結合処理のメソッド呼び出す
                //         Item combinePDF = Inn.applyMethod("z_dcn_pmd_pdf_combine", combineParam);
                //         if (!combinePDF.getProperty("result", "false").Equals("true"))
                //             throw new Exception("DCN + PMD PDFマージに失敗しました。");

                //         finalPdfPath = combinePDF.getProperty("result_json", "").Trim('[', ']', '"').Trim();
                //         if (!File.Exists(finalPdfPath))
                //             throw new Exception("結合後のPDFファイルが存在しません: " + finalPdfPath);

                //         batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "DCN + PMD PDFマージ完了");
                            
                //     }

                //     try
                //     {
                //         StampPdf(finalPdfPath, DCN.getProperty("keyed_name"));
                //     }
                //     catch (Exception ex)
                //     {
                //         batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, "スタンプ処理失敗: " + ex.Message);
                //     }

                //     string uploadDir = Path.Combine(Path.GetTempPath(), "z_dcn_upload");
                //     if (!Directory.Exists(uploadDir))
                //         Directory.CreateDirectory(uploadDir);

                //     string uploadFilePath = Path.Combine(uploadDir, Path.GetFileName(finalPdfPath));
                //     File.Copy(finalPdfPath, uploadFilePath, true);

                //     Item uploadItem = Inn.newItem("z_upload_download_file", Actions.ADD);
                //     uploadItem.setProperty("z_file_name", pmdFlag.Equals("Yes", StringComparison.OrdinalIgnoreCase) ? "DCN_PMD_Sheet_Merged.pdf" : ExportFile.Name);
                //     uploadItem.setFileProperty("z_file", uploadFilePath);
                //     uploadItem = uploadItem.apply();
                //     if (uploadItem.isError())
                //        throw new Exception(uploadItem.getErrorString());

                //     // DCN に添付
                //     DCN.setAction(Actions.EDIT);
                //     DCN.setAttribute("serverEvents", "0");
                //     DCN.setProperty("z_dcn_file", uploadItem.getProperty("z_file", ""));
                //     DCN = DCN.apply();
                //     if (DCN.isError())
                //         throw new Exception(DCN.getErrorString());

                //     DeleteWorkDirectory();
                //     batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "PDF処理完了");

                //     return this;
                //     // SDBMS_3.3.1 維持10月 PMD シート追加 RGK ポーイピュ END

                // }
                // アクション実行の場合、ファイルアイテムに登録
                else
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "Upload Download Fileに保存");
                    // 返却アイテム
                    Item resultItem = Inn.newItem("z_upload_download_file");
                    resultItem.setAction(Actions.ADD);
                    resultItem.setProperty("z_file_name", ExportFile.Name);
                    resultItem.setFileProperty("z_file", ExportFile.FullName);
                    resultItem = resultItem.apply();

                    //作業フォルダの削除
                    DeleteWorkDirectory();
                    return resultItem;
                }
                
            }
            catch (Aras.Server.Core.InnovatorServerException aex)
            {
                XmlDocument errDom = new XmlDocument();
                aex.ToSoapFault(errDom);
                errMsg = errDom.GetElementsByTagName("faultstring")[0].InnerText;

                if (batchLogger != null)
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, errMsg);
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, aex.ToString());
                }
                else
                {
                    RGKCustomize.CommonTools.AppLog(Inn, ThisCCO, METHOD_NAME, "System Error : ", aex.ToString() + targetinfo);
                }
                return Inn.newError(errMsg);
            }
            catch (Exception ex)
            {
                if (batchLogger != null)
                {
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, errMsg);
                    batchLogger.WriteLog(BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.INFO, ex.ToString());
                }
                else
                {
                    RGKCustomize.CommonTools.AppLog(Inn, ThisCCO, METHOD_NAME, "System Error : ", ex.ToString() + targetinfo);
                }

                // 処理結果の更新

                // アクションの場合はErrorを返す
                return Inn.newError(ex.Message);

                // WFの場合はエラーを返さず終了
            }
            finally
            {
                // 作業ファイルが残っている場合は解放
                if (ExportExcel != null) ExportExcel.Dispose();
                // ログ出力終了
                if (batchLogger != null) batchLogger.Dispose();
            }
            return this;

        }

        // SDBMS_3.3.1 維持10月 PMD シート追加 RGK ポーイピュ START
        // Attachmentsタブからpdfを取得
        private Item GetTargetItemVaultList(Item item)
        {
            Item retList = Inn.newItem();
            var itemId = item.getID();
            var itemTypeId = item.getAttribute("typeId");

            var itemRelTypeList = Inn.newItem("RelationshipType", "get");
            itemRelTypeList.setProperty("source_id", itemTypeId);

            var itemFile = Inn.newItem("ItemType", "get");
            itemFile.setProperty("name", "File");
            itemRelTypeList.setRelatedItem(itemFile);
            itemRelTypeList = itemRelTypeList.apply();

            for (int i = 0; i < itemRelTypeList.getItemCount(); i++)
            {
                var relType = itemRelTypeList.getItemByIndex(i);
                var relName = relType.getProperty("name");
                var rel = Inn.newItem(relName, "get");
                rel.setProperty("source_id", itemId);
                rel = rel.apply();

                if (!rel.isEmpty())
                {
                    retList.appendItem(rel);
                }
            }
            return retList;
        }

        // Attachmentsタブから取得したPDFをダウンロード
        private PdfFileInfoClass DownloadPdfFile(string downloadFolderPath, Item fileItem)
        {
            if (fileItem == null)
            {
                return null;
            }
            // ファイル名"PMD Sheet.pdf"のみダウンロード
            string fileName = fileItem.getProperty("filename", "");
            if (!fileName.Equals("PMD Sheet.pdf", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            string fileFolder = Path.Combine(downloadFolderPath, fileItem.getID());
            System.IO.Directory.CreateDirectory(fileFolder);

            fileItem.checkout(fileFolder);
            string filePath = Path.Combine(fileFolder, fileName);
            var ret = new PdfFileInfoClass
            {
                FileFolder = fileFolder,
                FilePath = filePath,
                FileName = fileName
            };
            return ret;
        }

        private void StampPdf(string pdfPath, string dcnNo)
        {
            using (PtlPDFDocument pdfDoc = new PtlPDFDocument())
            {
                var inputFile = new PtlParamInput(pdfPath);
                pdfDoc.load(inputFile);

                using (PtlPages pages = pdfDoc.getPages())
                {
                    int pageCount = pages.getCount();
                    for (int i = 0; i < pageCount; i++)
                    {
                        using (PtlPage page = pages.get(i))
                        {
                            // DCN番号、ページ数付けメソッド
                            WriteString(page, pageCount, i + 1, dcnNo);
                        }
                    }
                }

                var outputFile = new PtlParamOutput(pdfPath);
                pdfDoc.save(outputFile);
                inputFile.Dispose();
            }
        }

        // DCN番号、ページ数付けメソッド
        private void WriteString(PtlPage page, int pagecount, int nowpage, string dcnno)
        {
            using (PtlContent content = page.getContent())
            {
                using (PtlParamWriteString writestring = new PtlParamWriteString())
                using (PtlParamFont font = new PtlParamFont())
                {
                    // フォント設定
                    font.setName("Arial");
                    font.setSize(10.0f);
                    writestring.setFont(font);

                    // 文字色設定
                    using (PtlColorDeviceRGB colorText = new PtlColorDeviceRGB(0.0f, 0.0f, 0.0f))
                    {
                        writestring.setTextColor(colorText);
                    }

                    using (PtlRect rectPage = page.getMediaBox())
                    {
                        float width = rectPage.getRight() - rectPage.getLeft();
                        float height = rectPage.getTop() - rectPage.getBottom();                        
                        float topMargin = 37.5f;
                        float rightMargin = 10.0f;

                        using (PtlRect rectPageNum = new PtlRect(
                            width - rightMargin - 150.0f,
                            height - topMargin,
                            width - rightMargin,
                            height - topMargin + 15.0f
                        ))
                        {
                            content.writeString(rectPageNum, PtlContent.ALIGN.ALIGN_TOP_RIGHT, dcnno + "   (" + nowpage + "/" + pagecount + ") ", writestring);
                        }
                    }
                }
            }
        }
        // SDBMS_3.3.1 SDBMS_3.3.1 維持10月 PMD シート追加 RGK ポーイピュ END

        private void InitDefinitions()
        {
            // 処理結果アイテムを取得

            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "帳票定義アイテムを取得");
            // 帳票定義アイテムを取得
            ReportDefinition = Inn.getItemByKeyedName(Items.Report.TYPE_NAME, "z_dcn_definition");
            // シート定義および出力定義を取得
            ReportDefinition.fetchRelationships(Items.Report.SheetDef.REL_NAME);
            ReportDefinition.fetchRelationships(Items.Report.ExportDef.REL_NAME);
            SheetDefinitions = ReportDefinition.getRelationships(Items.Report.SheetDef.REL_NAME);
            ExportDefinition = ReportDefinition.getRelationships(Items.Report.ExportDef.REL_NAME);
            // ページカウントを保持する変数を初期化
            int sheetDefCount = SheetDefinitions.getItemCount();
            PageCountList = new int[sheetDefCount];
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "バッチ設定を取得");
            // バッチ設定を取得
            BatchSetting = Inn.getItemByKeyedName("z_batch_setting", "z_dcn_export");
            BatchSetting.fetchRelationships("z_rel_batch_setting_process");
            BatchSettingProcess = BatchSetting.getRelationships("z_rel_batch_setting_process").getItemByIndex(0);
            TempFolderPath = BatchSettingProcess.getProperty("z_output_path", "");
        }

        /// <summary>
        /// DCN情報の出力
        /// </summary>
        private void OutputDCNInfo()
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "DCN情報出力");

            // 出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(ExportSheetIndexDef.DCN, ExportBlockDef.BASIC_INFO);

            int defCount = exportDefs.getItemCount();
            int fromRow = 0;
            int fromCol = 0;
            int toRow = 0;
            int toCol = 0;
            int totalBreakCount = 0;
            int refRow = int.Parse(exportDefs.getItemByIndex(0).getProperty("z_row", "1"));
            for (int i = 0; i < defCount; i++)
            {
                // 定義を取得
                Item exportDef = exportDefs.getItemByIndex(i);
                string isExport = exportDef.getProperty("z_is_export", "0");
                if (isExport == "0") continue;
                string propName = exportDef.getProperty("z_column_name", "");
                string label = exportDef.getProperty("z_label", "");
                fromRow = int.Parse(exportDef.getProperty("z_row", "1"));
                fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                string needShrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                string wrapText = exportDef.getProperty("z_wrap_text", "0");
                string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                // 出力定義行数
                int tempRowNum = int.Parse(exportDef.getProperty("z_row_merge", "1"));
                int outputRowNum = tempRowNum;
                // 必要な加算行数
                int addRowNum = 0;
                // 出力値取得
                if (string.IsNullOrEmpty(propName))
                {
                    // ラベル出力
                    ExportSheet.Cells[fromRow, fromCol].Value = label;
                    ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    // 値は出力しない
                    continue;
                }
                string value = DCN.getProperty(propName.ToLower(), "");
                string[] valueLines = value.Split("\n");

                // 1.折り返して表示するプロパティの表示に必要な行数を取得
                if (wrapText.Equals("1"))
                {
                    // 1-1.改行コードの数を確認

                    int needRowNum = valueLines.Length;

                    // 1-2.改行コードまでの間に自動改行が発生するか文字数で確認
                    int breakRowInLine = 0;
                    foreach (string line in valueLines)
                    {
                        // 表示に必要な行数を取得
                        breakRowInLine = GetBreakRows(53, line);

                        // 1-3.自動改行が発生する数を加算
                        if (0 < breakRowInLine)
                        {
                            // 1行以上必要な分を加算
                            needRowNum = needRowNum + (breakRowInLine - 1);
                        }
                    }

                    // 1-5.定義の行数に追加が必要な行数を取得
                    if (tempRowNum < needRowNum)
                    {
                        addRowNum = needRowNum - tempRowNum;
                    }

                    // 1-6.出力する行数を取得
                    outputRowNum = tempRowNum + addRowNum;

                }

                // 2. ラベルと値出力の間に改ページが発生しないか確認
                int chkFromRow = CurrentPosition.Y + totalBreakCount + (int.Parse(exportDef.getProperty("z_row", "1")) - refRow);
                int chkToRow = chkFromRow + (outputRowNum - 1);

                if (IsBreakRow(chkToRow + 1))
                {
                    // 折り返して表示するセルの場合、出力行数が1ページに収まるか確認
                    bool canFitPage = true;
                    if (wrapText.Equals("1"))
                    {
                        int maxRowCount = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_MAX_DETAIL_ROWS, "0"));
                        canFitPage = outputRowNum < maxRowCount;
                    }
                    // １ページに収まる場合は改ページして出力
                    if (canFitPage)
                    {
                        SetCurrentPosition(1, chkToRow + 1);
                        // 改ページしたら合計改行数をリセット
                        totalBreakCount = 0;
                        // 改ページしたら基準ポジションをリセット
                        refRow = int.Parse(exportDef.getProperty("z_row", "1"));
                    }
                    // 1ページに収まらない場合はエラーとする(※要望があれば分けて表示)
                    else
                    {
                        throw new Exception("1つの属性値が1ページの印字範囲を超えています。");
                    }
                }

                // 3.出力場所を算出
                // 基本情報出力開始行 + 改行加算の合計数 + 基本情報出力開始行からの定義行数
                fromRow = CurrentPosition.Y + totalBreakCount + (int.Parse(exportDef.getProperty("z_row", "1")) - refRow);
                fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                toRow = fromRow + (outputRowNum - 1);
                toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);

                // 3.加算した行数の合計数を保持
                totalBreakCount = totalBreakCount + addRowNum;

                // ラベル出力
                ExportSheet.Cells[fromRow, fromCol].Value = label;
                ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;

                // セル位置を操作
                fromRow++;
                toRow++;
                switch (propName)
                {
                    case "z_type_code":
                        if (!string.IsNullOrEmpty(value)) value = Inn.getItemById("z_type_code", value).getProperty("z_type_code", "");
                        break;

                    case "z_ata_chapter_no":
                        if (!string.IsNullOrEmpty(value)) value = Inn.getItemById("z_ata_chapter_no", value).getProperty("z_ata_chapter_no", "");
                        break;

                    case "z_affiliation_code":
                        if (!string.IsNullOrEmpty(value)) value = Inn.getItemById("z_affiliation_code", value).getProperty("z_affiliation_code", "");
                        ExportSheet.Cells[fromRow, fromCol].Value = value;
                        break;

                    case "z_assigned_creator":
                        if (!string.IsNullOrEmpty(value)) value = GetIdentityName(value);
                        break;

                    case "z_designated_user":
                        if (!string.IsNullOrEmpty(value)) value = GetIdentityName(value);
                        break;

                    case "z_rev":
                        value = newPart.getProperty("major_rev", "");
                        break;


                    default:
                        break;
                }
                // 日英混在プロパティはフォントを変更
                if (EnJaProps.Contains(propName))
                {
                    SetValueWithFontName(value, fromRow, fromCol);
                }
                else
                {
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                }

                // 折り返して表示するプロパティ用の書式を設定
                if (wrapText.Equals("1"))
                {
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;
                    ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                }
                // それ以外用の書式を設定
                else
                {
                    ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                if (needShrinkToFit == "1") { ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true; }

            }

            // 操作セルを現在の出力位置にもってくる
            SetCurrentPosition(1, toRow);

            // 次の項目の間に空ける行数
            int brankRow = 1;
            for (int i = 0; i < brankRow + 1; i++)
            {
                // 操作セルを次に進める
                if (IsBreakRow(CurrentPosition.Y + 1))
                {
                    SetCurrentPosition(1, ++CurrentPosition.Y);
                    // 改ページなら終了
                    break;
                }
                SetCurrentPosition(1, ++CurrentPosition.Y);
            }



            return;
        }

        /// <summary>
        /// 必要行数の取得
        /// </summary>
        /// <param name="charNum"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private int GetBreakRows(int charNum, string value)
        {
            int breakRow = 0;
            int valueLength = value.Length;

            breakRow = valueLength / charNum;

            // 端数繰り上げ
            if (0 < (valueLength % charNum))
            {
                breakRow++;
            }

            return breakRow;
        }



        /// <summary>
        /// 履歴の出力
        /// </summary>
        /// <param name="drsID"></param>
        private void OutputHistory(string drsID)
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "History出力");

            // 出力データ取得 TODO
            Item sqlRet = CallProcedure("z_get_drs_dcn_approval_history", drsID);

            int dataCount = sqlRet.getItemCount();
            if (dataCount == 0)
            {
                return;
            }

            // ラベルの出力
            outputLabel(1, ExportSheetIndexDef.DCN, ExportBlockDef.HISTORY_LABEL, false);

            // ラベル出力数分下へ移動
            SetCurrentPosition(1, ++CurrentPosition.Y);

            // 値の出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(ExportSheetIndexDef.DCN, ExportBlockDef.HISTORY_VALUE);
            int defCount = exportDefs.getItemCount();
            for (int d = 0; d < dataCount; d++)
            {
                bool isSkip = false;
                Item retRow = sqlRet.getItemByIndex(d);
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string isExport = exportDef.getProperty("z_is_export", "0");
                    if (isExport == "0") continue;
                    string propName = exportDef.getProperty("z_column_name", "");
                    int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                    int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                    int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                    int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    string needShrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");

                    //出力値取得
                    string value = retRow.getProperty(propName.ToLower(), "");
                    List<string> showValues;

                    switch (propName)
                    {
                        case "Activity":
                            switch (fromCol)
                            {
                                case 3:
                                    if (Activities.TryGetValue(key: value, value: out showValues) == true)
                                    {

                                        value = showValues[0];
                                    }
                                    else
                                    {
                                        isSkip = true;
                                        continue;
                                    }
                                    break;

                                case 8:
                                    if (Activities.TryGetValue(key: value, value: out showValues) == true)
                                    {
                                        value = showValues[1];
                                    }
                                    else
                                    {
                                        isSkip = true;
                                        continue;
                                    }
                                    break;

                                default: break;
                            }
                            break;

                        case "Identity":
                            break;

                        case "z_proposal_division0":
                            break;

                        default:

                            break;

                    }

                    if (isSkip)
                    {
                        break;
                    }

                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (needShrinkToFit == "1") { ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true; }

                }

                if (isSkip) continue;

                // 改ページする場合はヘッダを再出力
                if (d + 1 != dataCount)
                {
                    if (SetCurrentPosition(1, ++CurrentPosition.Y))
                    {
                        outputLabel(1, ExportSheetIndexDef.DCN, ExportBlockDef.HISTORY_LABEL, false);
                        // ラベル出力数分下へ移動
                        SetCurrentPosition(1, ++CurrentPosition.Y);
                    }
                }
            }
        }

        /// <summary>
        /// Part基本情報の差分出力
        /// 基準点からの相対位置(帳票定義の行)に出力する
        /// </summary>
        /// <param name="partID"></param>
        /// <param name="sheetDef"></param>
        /// <param name="labelDef"></param>
        /// <param name="valueDef"></param>
        private void OutputPartDiffInfoVariable(string sheetDef, string labelDef, string valueDef, Item outputDataList)
        {
            int dataCount = outputDataList.getItemCount();

            // 出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetDef, valueDef);
            int defCount = exportDefs.getItemCount();

            // 出力位置を初期化
            int fromRow = 0;
            int fromCol = 0;
            int toRow = 0;
            int toCol = 0;
            //int totalBreakCount = 0;
            int bfDefRow = 0;
            int currentRow = CurrentPosition.Y;
            //int refRow = int.Parse(exportDefs.getItemByIndex(0).getProperty("z_row", "1"));
            for (int i = 0; i < defCount; i++)
            {


                // 出力定義を1行ずつ処理
                Item exportDef = exportDefs.getItemByIndex(i);

                // 行の可変が必要な項目か
                bool isVariable = exportDef.getProperty("z_wrap_text", "0").Equals("1");

                // 定義を取得
                string isExport = exportDef.getProperty("z_is_export", "0");
                if (isExport == "0") continue;
                string label = exportDef.getProperty("z_label", "");
                string propName = exportDef.getProperty("z_column_name", "");
                string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");
                int defRow = int.Parse(exportDef.getProperty("z_row", "1"));
                if (bfDefRow == 0)
                {
                    bfDefRow = defRow;
                }
                // 前の行と定義の行が異なる場合は行を進める
                if (bfDefRow == defRow)
                {
                    currentRow = CurrentPosition.Y;
                }
                else
                {
                    SetCurrentPosition(1, currentRow);
                }
                bfDefRow = defRow;

                // ラベル + IS + WASの出力で必要な行数を求める

                // ラベル行数
                int labelRow = 1;
                // 出力定義行数
                int tempRowNum = int.Parse(exportDef.getProperty("z_row_merge", "1"));
                // 値に必要な行数(IS/WAS)
                List<int> outputRowNums = new List<int>();

                // 行可変の場合は出力するために必要な領域を算出
                int outputRow = 0;
                if (isVariable)
                {
                    for (int d = 0; d < dataCount; d++)
                    {
                        // 値を取得
                        Item dataRow = outputDataList.getItemByIndex(d);
                        string value = dataRow.getProperty(propName.ToLower(), "");
                        // 必要行数を取得
                        int outputRowNum = GetOutputNeedRow(tempRowNum, 53, value);
                        // IS/WASの順に保持
                        outputRowNums.Add(outputRowNum);
                    }
                }

                // 出力に必要な領域を算出
                if (isVariable)
                {
                    outputRow = labelRow;
                    foreach (int row in outputRowNums)
                    {
                        outputRow += row;
                    }
                    // 必要な領域が1ページを超えていないか確認
                    bool canFitPage = true;
                    if (isVariable)
                    {
                        int maxRowCount = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_MAX_DETAIL_ROWS, "0"));
                        if (maxRowCount < outputRow)
                        {
                            throw new Exception(label + "の出力が1ページの印字範囲を超えています。");
                        }
                    }
                }
                else
                {
                    outputRow = labelRow + tempRowNum * dataCount;
                }

                int chkToRow = currentRow + (outputRow - 1);
                if (IsBreakRow(chkToRow))
                {
                    // 改ページ
                    SetCurrentPosition(1, chkToRow);
                    currentRow = CurrentPosition.Y;
                }

                //ラベルの出力
                fromRow = currentRow;
                fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                outputLabelVariable(fromRow, fromCol, exportDef, false);
                currentRow++;

                for (int d = 0; d < dataCount; d++)
                {
                    fromRow = currentRow;

                    if (isVariable)
                    {
                        toRow = fromRow + outputRowNums[d] - 1;
                        // 加算合計数を保持
                    }
                    else
                    {
                        toRow = fromRow;
                    }
                    toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    //batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, propName + "[" + fromRow + "," + fromCol + "," + toRow + "," + toCol + "]");

                    // IS-WAS出力
                    ExportSheet.Cells[fromRow, 3].Style.Font.Size = 10;
                    int rowNum = d % 2;
                    string isWas = "";
                    if (rowNum == 0)
                    {
                        isWas = "(IS)";
                    }
                    else
                    {
                        isWas = "(WAS)";
                    }
                    ExportSheet.Cells[fromRow, 3].Value = isWas;
                    ExportSheet.Cells[fromRow, 3, fromRow, 4].Merge = true;
                    ExportSheet.Cells[fromRow, 3, fromRow, 4].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                    ExportSheet.Cells[fromRow, 3, fromRow, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // 値を取得
                    Item dataRow = outputDataList.getItemByIndex(d);
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    if (isVariable) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    else ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (isVariable) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                    // 操作位置を進める
                    currentRow = toRow + 1;
                }
            }

            // 操作セルを現在の出力位置にもってくる
            SetCurrentPosition(1, currentRow);

            // 次の項目の間に空ける行数
            int brankRow = 0;
            for (int i = 0; i < brankRow + 1; i++)
            {
                // 操作セルを次に進める
                if (IsBreakRow(CurrentPosition.Y + 1))
                {
                    SetCurrentPosition(1, ++CurrentPosition.Y);
                    // 改ページなら終了
                    break;
                }
                SetCurrentPosition(1, ++CurrentPosition.Y);
            }

            return;
        }

        /// <summary>
        /// 値の出力に必要な行数を取得
        /// </summary>
        /// <param name="tempRowNum"></param>
        /// <param name="charNum"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private int GetOutputNeedRow(int tempRowNum, int charNum, string value)
        {
            int outputRowNum = 1;
            // 必要な加算行数
            int addRowNum = 0;

            // 1-1.改行コードの数を確認
            string[] valueLines = value.Split("\n");
            int needRowNum = valueLines.Length;

            // 1-2.改行コードまでの間に自動改行が発生するか文字数で確認
            int breakRowInLine = 0;
            foreach (string line in valueLines)
            {
                // 表示に必要な行数を取得
                breakRowInLine = GetBreakRows(53, line);

                // 1-3.自動改行が発生する数を加算
                if (0 < breakRowInLine)
                {
                    // 1行以上必要な分を加算
                    needRowNum = needRowNum + (breakRowInLine - 1);
                }
            }

            // 1-4.定義の行数に追加が必要な行数を取得
            if (tempRowNum < needRowNum)
            {
                addRowNum = needRowNum - tempRowNum;
            }

            // 1-5.出力する行数を取得
            outputRowNum = tempRowNum + addRowNum;

            return outputRowNum;
        }

        /// <summary>
        /// Part基本情報の差分出力（行可変対応修正）
        /// </summary>
        /// <param name="sheetDef"></param>
        /// <param name="labelDef"></param>
        /// <param name="valueDef"></param>
        /// <param name="outputDataList"></param>
        private void OutputPartDiffInfo(string sheetDef, string labelDef, string valueDef, Item outputDataList)
        {
            // ========== ラベルの出力 ==========

            // 出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetDef, labelDef);
            int defCount = exportDefs.getItemCount();

            int fromRow = 0;
            int fromCol = 0;
            int toRow = 0;
            int toCol = 0;

            // 出力定義をフィルタ
            exportDefs = GetExportDefinitionsAt(sheetDef, valueDef);
            defCount = exportDefs.getItemCount();

            // データ行数
            int dataCount = outputDataList.getItemCount();

            int bfDefRow = 0;
            for (int i = 0; i < defCount; i++)
            {
                // 定義を取得
                Item exportDef = exportDefs.getItemByIndex(i);
                string isExport = exportDef.getProperty("z_is_export", "0");
                if (isExport == "0") continue;
                string propName = exportDef.getProperty("z_column_name", "");
                string labelbase = exportDef.getProperty("z_label", "");
                string label = labelbase.Replace("$", NewLineChars.LF);
                string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                string wrapText = exportDef.getProperty("z_wrap_text", "0");
                string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");
                int defRow = int.Parse(exportDef.getProperty("z_row", "1"));
                if (bfDefRow == 0)
                {
                    bfDefRow = defRow;
                }
                // 前の行と定義の行が異なる場合は行を進める
                if (bfDefRow != defRow)
                {
                    SetCurrentPosition(1, CurrentPosition.Y + 3);
                }
                bfDefRow = defRow;

                // 出力位置を計算
                fromRow = CurrentPosition.Y;
                // 改ページ確認
                if (IsBreakRow(fromRow + 2))
                {
                    SetCurrentPosition(1, fromRow + 2);
                    fromRow = CurrentPosition.Y;
                }
                fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);


                // ラベル出力
                // フォントサイズ指定
                if (labelbase.Equals("Serial Number Control") || labelbase.Equals("Line Replaceable Unit") ||
                    labelbase.Equals("Time$Controlled Item") || labelbase.Equals("Serial Number$Control") || labelbase.Equals("Line Replaceable$Unit"))
                {
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 9;
                }
                else
                {
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                }
                ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                if (!string.IsNullOrEmpty(label)) ExportSheet.Cells[fromRow, fromCol].Value = label;
                // IS-WAS出力
                ExportSheet.Cells[fromRow + 1, fromCol - 2].Value = "(IS)";
                ExportSheet.Cells[fromRow + 1, fromCol - 2, toRow + 1, fromCol - 1].Merge = true;
                ExportSheet.Cells[fromRow + 1, fromCol - 2, toRow + 1, fromCol - 1].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                ExportSheet.Cells[fromRow + 2, fromCol - 2].Value = "(WAS)";
                ExportSheet.Cells[fromRow + 2, fromCol - 2, toRow + 2, fromCol - 1].Merge = true;
                ExportSheet.Cells[fromRow + 2, fromCol - 2, toRow + 2, fromCol - 1].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");

                // データ行数分
                for (int d = 0; d < dataCount; d++)
                {
                    Item dataRow = outputDataList.getItemByIndex(d);
                    string value = dataRow.getProperty(propName.ToLower(), "");

                    fromRow++;
                    toRow++;

                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (wrapText.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                }

            }

            // 操作セルを次の行に進める
            SetCurrentPosition(1, CurrentPosition.Y + 3);

            return;
        }


        /// <summary>
        /// ラベルの出力
        /// </summary>
        /// <param name="headCount"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="labelBlock"></param>
        /// <param name="isMerge"></param>
        private void outputLabel(int headCount, string sheetIndex, string labelBlock, bool isMerge)
        {
            // ラベルの出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetIndex, labelBlock);
            int defCount = exportDefs.getItemCount();

            // ヘッダを出力することで印字可能域を超過する、またはヘッダのみ出力される場合は改ページ
            for (int i = 1; i <= headCount; i++)
            {
                if (IsBreakRow(CurrentPosition.Y + i))
                {
                    SetCurrentPosition(1, CurrentPosition.Y + i);
                    break;
                }
            }

            for (int i = 0; i < defCount; i++)
            {
                // 定義を取得
                Item exportDef = exportDefs.getItemByIndex(i);
                string isExport = exportDef.getProperty("z_is_export", "0");
                if (isExport == "0") continue;
                string propName = exportDef.getProperty("z_column_name", "");
                string labelbase = exportDef.getProperty("z_label", "");
                string label = labelbase.Replace("$", NewLineChars.LF);
                int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                string needShrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                // フォントサイズ指定
                if (labelbase.Equals("Serial Number Control") || labelbase.Equals("Line Replaceable Unit") ||
                    labelbase.Equals("Time$Controlled Item") || labelbase.Equals("Serial Number$Control") || labelbase.Equals("Line Replaceable$Unit"))
                {
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 9;
                }
                else
                {
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                }

                // ラベル出力
                if (!string.IsNullOrEmpty(label)) ExportSheet.Cells[fromRow, fromCol].Value = label;
                ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = isMerge;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = isMerge;
                if (isMerge)
                {
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;
                }

            }

            return;
        }

        private void outputLabelVariable(int fromRow, int fromCol, Item exportDef, bool isMerge)
        {
            // 定義を取得
            string isExport = exportDef.getProperty("z_is_export", "0");
            if (isExport == "0") return;
            string labelbase = exportDef.getProperty("z_label", "");
            if (string.IsNullOrEmpty(labelbase)) return;
            string label = labelbase.Replace("$", NewLineChars.LF);
            string propName = exportDef.getProperty("z_column_name", "");
            int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
            int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
            string needShrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
            string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

            // 書式設定
            ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
            //batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, label + "[" + fromRow + "," + fromCol + "," + toRow + "," + toCol + "]");
            ExportSheet.Cells[fromRow, fromCol].Value = label;
            ExportSheet.Cells[fromRow, fromCol].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = isMerge;
            ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = isMerge;
            if (isMerge)
            {
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;
            }
            return;
        }


        /// <summary>
        /// Noteの出力
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="labelBlock"></param>
        /// <param name="valueBlock"></param>
        /// <param name="outputDataList"></param>
        private void OutputNoteDiff(string sheetIndex, string labelBlock, string valueBlock, Item outputDataList, bool isMerge)
        {
            // 値の出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetIndex, valueBlock);
            int defCount = exportDefs.getItemCount();

            // 出力判断
            bool isSkip = true;
            // データ行数分確認
            int dataCount = outputDataList.getItemCount();
            int d = 0;
            while (isSkip && d < dataCount)
            {
                Item dataRow = outputDataList.getItemByIndex(d);
                // 定義の行数分確認
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string propName = exportDef.getProperty("z_column_name", "");
                    // 出力値取得
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    // 出力値が見つかったら出力すると判断
                    if (!string.IsNullOrEmpty(value))
                    {
                        isSkip = false;
                        break;
                    }
                }
                d++;
            }

            // 出力値がない場合は終了
            if (isSkip) return;

            // ラベルの出力
            outputLabel(1, sheetIndex, labelBlock, isMerge);

            // 操作中セルを設定
            CurrentPosition.Y++;

            // 値の出力 データ行数分
            for (d = 0; d < dataCount; d++)
            {
                Item dataRow = outputDataList.getItemByIndex(d);

                // 出力するために必要な領域を算出

                // 値に必要な行数(IS/WAS)
                List<int> outputRowNums = new List<int>();
                int outputRow = 0;
                // 値を取得
                string value = dataRow.getProperty("z_note", "");
                // 必要行数を取得
                outputRow = GetOutputNeedRow(1, 45, value);

                int chkToRow = CurrentPosition.Y + (outputRow - 1);
                if (IsBreakRow(chkToRow))
                {
                    // 改ページ
                    SetCurrentPosition(1, chkToRow + 1);
                    // ラベルの出力
                    outputLabel(1, sheetIndex, labelBlock, isMerge);
                    // 操作中セルを設定
                    CurrentPosition.Y++;
                }

                // IS-WAS出力
                ExportSheet.Cells[CurrentPosition.Y, 3].Style.Font.Size = 10;
                int rowNum = d % 2;
                string isWas = "";
                if (rowNum == 0)
                {
                    isWas = "(IS)";
                }
                else
                {
                    isWas = "(WAS)";
                }
                ExportSheet.Cells[CurrentPosition.Y, 3].Value = isWas;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Merge = true;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // 定義の行数分
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string isExport = exportDef.getProperty("z_is_export", "0");
                    if (isExport == "0") continue;
                    string propName = exportDef.getProperty("z_column_name", "");
                    //int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                    int fromRow = CurrentPosition.Y;
                    int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                    int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1) + (outputRow - 1);
                    int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                    string wrapText = exportDef.getProperty("z_wrap_text", "0");
                    string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                    //出力値取得
                    value = dataRow.getProperty(propName.ToLower(), "");
                    if (int.Parse(exportDef.getProperty("z_col_merge", "1")) == 1)
                    {
                        isMerge = false;
                    }
                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (wrapText.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                }

                // 次の行に進める
                //if (d + 1 != dataCount && SetCurrentPosition(1, CurrentPosition.Y + outputRow))
                //{
                //    // 改ページする場合はヘッダを再出力
                //    outputLabel(2, sheetIndex, labelBlock, isMerge);
                //}
                if (SetCurrentPosition(1, CurrentPosition.Y + outputRow))
                {
                    // 改ページする場合はヘッダを再出力
                    outputLabel(2, sheetIndex, labelBlock, isMerge);
                }
            }

            // 次の行に進める
            //SetCurrentPosition(1, ++CurrentPosition.Y);

        }

        /// <summary>
        /// SPECの出力
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="labelBlock"></param>
        /// <param name="valueBlock"></param>
        /// <param name="outputDataList"></param>
        private void OutputSpecDiff(string sheetIndex, string labelBlock, string valueBlock, Item outputDataList, bool isMerge)
        {
            // 値の出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetIndex, valueBlock);
            int defCount = exportDefs.getItemCount();

            // 出力判断
            bool isSkip = true;
            // データ行数分確認
            int dataCount = outputDataList.getItemCount();
            int d = 0;
            while (isSkip && d < dataCount)
            {
                Item dataRow = outputDataList.getItemByIndex(d);
                // 定義の行数分確認
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string propName = exportDef.getProperty("z_column_name", "");
                    // 出力値取得
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    // 出力値が見つかったら出力すると判断
                    if (!string.IsNullOrEmpty(value))
                    {
                        isSkip = false;
                        break;
                    }
                }
                d++;
            }

            // 出力値がない場合は終了
            if (isSkip) return;

            // ラベルの出力
            outputLabel(1, sheetIndex, labelBlock, isMerge);

            // 操作中セルを設定
            CurrentPosition.Y++;

            // 値の出力 データ行数分
            for (d = 0; d < dataCount; d++)
            {
                Item dataRow = outputDataList.getItemByIndex(d);

                // IS-WAS出力
                ExportSheet.Cells[CurrentPosition.Y, 3].Style.Font.Size = 10;
                int rowNum = d % 2;
                string isWas = "";
                if (rowNum == 0)
                {
                    isWas = "(IS)";
                }
                else
                {
                    isWas = "(WAS)";
                }
                ExportSheet.Cells[CurrentPosition.Y, 3].Value = isWas;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Merge = true;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // 定義の行数分
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string isExport = exportDef.getProperty("z_is_export", "0");
                    if (isExport == "0") continue;
                    string propName = exportDef.getProperty("z_column_name", "");
                    //int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                    int fromRow = CurrentPosition.Y;
                    int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                    int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                    int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                    string wrapText = exportDef.getProperty("z_wrap_text", "0");
                    string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                    //出力値取得
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    if (int.Parse(exportDef.getProperty("z_col_merge", "1")) == 1)
                    {
                        isMerge = false;
                    }
                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (wrapText.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                }

                // 次の行に進める
                if (d + 1 != dataCount && SetCurrentPosition(1, ++CurrentPosition.Y))
                {
                    // 改ページする場合はヘッダを再出力
                    outputLabel(2, sheetIndex, labelBlock, isMerge);
                }
            }

            // 次の行に進める
            SetCurrentPosition(1, ++CurrentPosition.Y);

        }

        /// <summary>
        /// 表形式で値を出力
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="labelBlock"></param>
        /// <param name="valueBlock"></param>
        /// <param name="outputDataList"></param>
        private void OutputNoteDiffVariable(string sheetIndex, string labelBlock, string valueBlock, Item outputDataList, bool isMerge)
        {
            // 値の出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetIndex, valueBlock);
            int defCount = exportDefs.getItemCount();

            // 出力判断
            bool isSkip = true;

            // データ行数分確認
            int dataCount = outputDataList.getItemCount();
            int d = 0;
            while (isSkip && d < dataCount)
            {
                Item chkDataRow = outputDataList.getItemByIndex(d);
                // 定義の行数分確認
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string propName = exportDef.getProperty("z_column_name", "");
                    // 出力値取得
                    string value = chkDataRow.getProperty(propName.ToLower(), "");
                    // 出力値が見つかったら出力すると判断
                    if (!string.IsNullOrEmpty(value))
                    {
                        isSkip = false;
                        break;
                    }
                }
                d++;
            }

            // 出力値がない場合は終了
            if (isSkip) return;

            // 1行目のNoteの出力に必要な行数取得
            d = 0;
            Item dataRow = outputDataList.getItemByIndex(d);
            int tempRowNum = 1;
            string note = dataRow.getProperty("z_note", "");
            int outputRowNum = GetOutputNeedRow(tempRowNum, 48, note);

            // ラベルの出力
            outputLabel(1 + outputRowNum, sheetIndex, labelBlock, isMerge);

            // 操作中セルを設定
            CurrentPosition.Y++;

            // 値の出力 データ行数分
            for (d = 0; d < dataCount; d++)
            {
                // IS-WAS出力
                ExportSheet.Cells[CurrentPosition.Y, 3].Style.Font.Size = 10;
                int rowNum = d % 2;
                string isWas = "";
                if (rowNum == 0)
                {
                    isWas = "(IS)";
                }
                else
                {
                    isWas = "(WAS)";
                }
                ExportSheet.Cells[CurrentPosition.Y, 3].Value = isWas;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Merge = true;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // 定義の行数分
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string isExport = exportDef.getProperty("z_is_export", "0");
                    if (isExport == "0") continue;
                    string propName = exportDef.getProperty("z_column_name", "");
                    //int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                    int fromRow = CurrentPosition.Y;
                    int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                    int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1) + (outputRowNum - 1);
                    int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                    string wrapText = exportDef.getProperty("z_wrap_text", "0");
                    string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                    //出力値取得
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    if (int.Parse(exportDef.getProperty("z_col_merge", "1")) == 1)
                    {
                        isMerge = false;
                    }
                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (wrapText.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                }

                // 次の行の出力行数確認
                if (d + 1 != dataCount)
                {
                    dataRow = getItemByIndex(d + 1);
                    note = dataRow.getProperty("z_note", "");
                    outputRowNum = GetOutputNeedRow(tempRowNum, 48, note);

                    if (SetCurrentPosition(1, ++CurrentPosition.Y))
                    {
                        // 改ページする場合はヘッダを再出力
                        outputLabel(1 + outputRowNum, sheetIndex, labelBlock, isMerge);
                    }
                }
            }

            // 次の行に進める
            SetCurrentPosition(1, ++CurrentPosition.Y);

        }

        /// <summary>
        /// 表形式で値を出力
        /// </summary>
        /// <param name="sheetIndex"></param>
        /// <param name="labelBlock"></param>
        /// <param name="valueBlock"></param>
        /// <param name="outputDataList"></param>
        private void OutputBOM(int headCount, string sheetIndex, string labelBlock, string valueBlock, Item outputDataList, bool isMerge, bool isEnd)
        {
            // 値の出力定義をフィルタ
            Item exportDefs = GetExportDefinitionsAt(sheetIndex, valueBlock);
            int defCount = exportDefs.getItemCount();

            // 出力判断
            bool isSkip = true;
            isSkip = IsSkipOutput(outputDataList, sheetIndex, valueBlock);

            //// 出力判断から除外するプロパティ
            //string[] skipProps = { "line_no", "sort_key", "item_number" };
            //// データ行数分確認
            //int dataCount = outputDataList.getItemCount();
            //int d = 0;
            //while (isSkip && d < dataCount)
            //{
            //    Item dataRow = outputDataList.getItemByIndex(d);
            //    // 定義の行数分確認
            //    for (int i = 0; i < defCount; i++)
            //    {
            //        // 定義を取得
            //        Item exportDef = exportDefs.getItemByIndex(i);
            //        string propName = exportDef.getProperty("z_column_name", "");
            //        // 出力判断から除外する場合は次へ
            //        if (skipProps.Contains(propName)) continue;
            //        // 出力値取得
            //        string value = dataRow.getProperty(propName.ToLower(), "");
            //        // 出力値が見つかったら出力すると判断
            //        if (!string.IsNullOrEmpty(value))
            //        {
            //            isSkip = false;
            //            break;
            //        }
            //    }
            //    d++;
            //}

            // 出力値がない場合は終了
            if (isSkip) return;
            // ラベルの出力
            outputLabel(headCount, sheetIndex, labelBlock, isMerge);
            // 操作中セルを設定
            SetCurrentPosition(1, CurrentPosition.Y + headCount);

            int dataCount = outputDataList.getItemCount();
            // 値の出力 データ行数分
            for (int d = 0; d < dataCount; d++)
            {
                Item dataRow = outputDataList.getItemByIndex(d);

                // IS-WAS出力
                ExportSheet.Cells[CurrentPosition.Y, 3].Style.Font.Size = 10;
                int rowNum = d % 2;
                string isWas = "";
                if (rowNum == 0)
                {
                    isWas = "(IS)";
                }
                else
                {
                    isWas = "(WAS)";
                }
                ExportSheet.Cells[CurrentPosition.Y, 3].Value = isWas;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Merge = true;
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 4].Style.HorizontalAlignment = GetExcelHorizontalAlignment("Center");
                ExportSheet.Cells[CurrentPosition.Y, 3, CurrentPosition.Y, 41].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // 定義の行数分
                for (int i = 0; i < defCount; i++)
                {
                    // 定義を取得
                    Item exportDef = exportDefs.getItemByIndex(i);
                    string isExport = exportDef.getProperty("z_is_export", "0");
                    if (isExport == "0") continue;
                    string propName = exportDef.getProperty("z_column_name", "");
                    //int fromRow = CurrentPosition.Y + int.Parse(exportDef.getProperty("z_row", "1")) - 1;
                    int fromRow = CurrentPosition.Y;
                    int fromCol = int.Parse(exportDef.getProperty("z_column", "1"));
                    int toRow = fromRow + (int.Parse(exportDef.getProperty("z_row_merge", "1")) - 1);
                    int toCol = fromCol + (int.Parse(exportDef.getProperty("z_col_merge", "1")) - 1);
                    string shrinkToFit = exportDef.getProperty("z_shrink_to_fit", "0");
                    string wrapText = exportDef.getProperty("z_wrap_text", "0");
                    string horizonAlignment = exportDef.getProperty("z_horizontal_alignment", "0");

                    // 出力値取得
                    string value = dataRow.getProperty(propName.ToLower(), "");
                    // ラベルは結合しない
                    if (int.Parse(exportDef.getProperty("z_col_merge", "1")) == 1)
                    {
                        isMerge = false;
                    }

                    // Line NO.は0埋めする
                    if (propName.Equals("line_no"))
                    {
                        value = value.PadLeft(3, '0');
                    }
                    if (propName.Equals("sort_key"))
                    {
                        value = value.Substring(0, value.LastIndexOf('_'));
                    }


                    // 値出力
                    ExportSheet.Cells[fromRow, fromCol].Style.Font.Size = 10;
                    ExportSheet.Cells[fromRow, fromCol].Value = value;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Merge = true;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = GetExcelHorizontalAlignment(horizonAlignment);
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    if (shrinkToFit.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.ShrinkToFit = true;
                    if (wrapText.Equals("1")) ExportSheet.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;

                }

                // 次の行に進める
                if (d + 1 != dataCount)
                {
                    if (SetCurrentPosition(1, ++CurrentPosition.Y))
                    {
                        // 改ページする場合はヘッダを再出力
                        outputLabel(headCount, sheetIndex, labelBlock, isMerge);
                        // 操作中セルを設定
                        SetCurrentPosition(1, CurrentPosition.Y + headCount);
                    }

                    // 次がISの場合はその次の行も確認
                    if (isWas.Equals("(WAS)"))
                    {
                        if (IsBreakRow(CurrentPosition.Y + 1))
                        {
                            // 改ページする場合はヘッダを再出力
                            outputLabel(headCount, sheetIndex, labelBlock, isMerge);
                            // 操作中セルを設定
                            SetCurrentPosition(1, CurrentPosition.Y + headCount);
                        }
                    }
                }
            }

            if (!isEnd)
            {
                SetCurrentPosition(1, ++CurrentPosition.Y);
                SetCurrentPosition(1, ++CurrentPosition.Y);
            }

        }

        /// <summary>
        /// 出力値を確認して、Skipするか判断する
        /// </summary>
        /// <param name="sheetIndexBlockList"></param>
        /// <param name="valueBlockList"></param>
        /// <param name="outputDataList"></param>
        /// <returns>true:Skipする false:Skipしない</returns>
        private bool IsSkipOutput(Item outputDataList, string sheetIndex, params string[] valueBlock)
        {
            bool isSkip = true;
            for (int p = 0; p < valueBlock.Length; p++)
            {
                // 値の出力定義をフィルタ
                Item exportDefs = GetExportDefinitionsAt(sheetIndex, valueBlock[p]);
                int defCount = exportDefs.getItemCount();

                // 出力判断から除外するプロパティ
                string[] skipProps = { "line_no", "sort_key", "item_number" };
                // データ行数分確認
                int dataCount = outputDataList.getItemCount();
                for (int d = 0; d < dataCount; d++)
                {
                    Item dataRow = outputDataList.getItemByIndex(d);
                    for (int i = 0; i < defCount; i++)
                    {
                        // 定義を取得
                        Item exportDef = exportDefs.getItemByIndex(i);
                        string propName = exportDef.getProperty("z_column_name", "");
                        // 出力判断から除外する場合は次へ
                        if (skipProps.Contains(propName)) continue;
                        // 出力値取得
                        string value = dataRow.getProperty(propName.ToLower(), "");
                        // 出力値が見つかったら出力すると判断
                        if (!string.IsNullOrEmpty(value))
                        {
                            return false;
                        }
                    }
                }
            }

            return isSkip;
        }


        /// <summary>
        /// ヘッダページ出力
        /// </summary>
        private void OutputPageNumber(int y, int sheetIndex, string dcnNo)
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "ページ数出力");

            // 定義を取得
            int hight = int.Parse(GetSheetDefinitionsAt(sheetIndex.ToString()).getProperty(Items.Report.SheetDef.Props.Z_SHEET_HEIGHT, "0"));
            string strTotalPageCount = TotalPageCount.ToString();

            // 1つ前のsheetIndex
            int beforeSheetIndex = (sheetIndex == 1 ? 0 : sheetIndex - 1);
            int pageCount = (sheetIndex == 1 ? 1 : PageCountList[beforeSheetIndex - 1] + 1);
            for (int i = 1; i <= PageCountList[sheetIndex - 1]; i++)
            {
                CurrentPageCount++;
                string strOutputPage = "(" + CurrentPageCount.ToString() + "/" + strTotalPageCount + ")";
                int currentRow = 3 + (i - 1) * hight;
                ExportSheet.Cells[currentRow, y].Value = strOutputPage;
                ExportSheet.Cells[currentRow, y].Style.Font.Size = 10;
                ExportSheet.Cells[currentRow, y, currentRow, y + 1].Merge = true;
                ExportSheet.Cells[currentRow, y, currentRow, y + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // DCN番号を出力
                ExportSheet.Cells[currentRow, y - 6].Value = dcnNo;
                ExportSheet.Cells[currentRow, y - 6].Style.Font.Size = 10;
                ExportSheet.Cells[currentRow, y - 6, currentRow, y - 1].Merge = true;
                ExportSheet.Cells[currentRow, y - 6, currentRow, y - 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            }

            return;
        }

        private void OutputLogo()
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "画像の出力");
            // 画像名、画像位置のDictionary
            Dictionary<string, int[]> pictureMap = new Dictionary<string, int[]>();
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "画像の出力1");
            ExcelWorksheet templateSheet = Worksheets[TemplateSheetIndex];
            ExcelDrawings drawings = templateSheet.Drawings;

            // 画像名順に取得されるため、画像の開始行でソートするようDictionaryに追加する
            foreach (ExcelDrawing drawing in drawings)
            {
                if (drawing.GetType() == typeof(ExcelPicture))
                {
                    ExcelPicture picture = (ExcelPicture)drawing;
                    // サイズが取得できるよう、画像形式を「TwoCellAnchor」にする
                    picture.ChangeCellAnchor(eEditAs.TwoCell);
                    int[] positions = new int[] { picture.From.Row, picture.From.Column, picture.To.Row };
                    pictureMap.Add(picture.Name, positions);
                    //リストに追加
                    excelPictures.Add(picture);
                }
            }
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "画像の出力2");


            foreach (ExcelPicture picture in excelPictures)
            {
                //検索文字列を取得
                string searchString = "template_image_" + picture.Name;

                //画像名で検索
                var query = from cell in ExportSheet.Cells[1, 1, ExportSheet.Dimension.End.Row, ExportSheet.Dimension.End.Column]
                            where cell.Text == searchString
                            select cell;
                if (0 < query.Count())
                {
                    int picNo = 0;
                    foreach (var rb in query.ToArray())
                    {
                        byte[] buffer = picture.Image.ImageBytes;

                        using (MemoryStream stream = new MemoryStream(buffer))
                        {
                            ExcelPicture addImage = ExportSheet.Drawings.AddPicture(picture.Name + picNo, stream);
                            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "画像の出力3");
                            // サイズが取得できるよう、出力する画像形式を「TwoCellAnchor」にする
                            addImage.ChangeCellAnchor(eEditAs.TwoCell);

                            // ポジションとサイズを設定
                            addImage.SetPosition(rb.Start.Row, 5, rb.Start.Column, 5);
                            var width = addImage.Size.Width;
                            addImage.SetSize((int)(double)width / 9525, (int)(double)addImage.Size.Height / 9525);
                        }
                        picNo++;

                        //テキストクリア
                        ExportSheet.Cells[rb.Start.Row, rb.Start.Column].Value = "";
                    }
                }
            }



            return;
        }


        /// <summary>
        /// 日本語と英語にフォント
        /// </summary>
        /// <param name="value"></param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        private void SetValueWithFontName(string value, int fromRow, int fromCol)
        {
            // セル内のリッチテキストを設定
            var cell = ExportSheet.Cells[fromRow, fromCol];
            var richText = cell.RichText;

            // 正規表現を使って、全角（日本語）部分と半角（英語）部分を分ける
            string[] textBlocks = System.Text.RegularExpressions.Regex.Split(value, @"([^\x01-\x7E\uFF61-\uFF9F]+)");

            foreach (var text in textBlocks)
            {
                if (string.IsNullOrEmpty(text)) continue;

                if (IsFullWidth(text)) // 全角部分
                {
                    richText.Add(text).FontName = "MS Pゴシック"; // 日本語部分にMS Gothicを設定
                }
                else if (string.IsNullOrWhiteSpace(text)) // 半角スペース部分
                {
                    richText.Add(text).FontName = "Arial"; // 半角スペースにArialを設定
                }
                else // 半角部分
                {

                    richText.Add(text).FontName = "Arial"; // 英語部分にArialを設定
                }
            }
        }

        /// <summary>
        /// 全角か半角かを判定
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private bool IsFullWidth(string input)
        {
            foreach (char c in input)
            {
                if (c >= 0x20 && c <= 0x7e) // 半角文字は範囲外
                {
                    return false; // 半角文字の場合はfalse
                }
            }
            return true; // 全角文字の場合はtrue
        }


        /// <summary>
        /// 印刷設定
        /// </summary>
        private void SetPrintSetting(int templateSheetIndex)
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "印刷設定");
            // シート幅取得
            int sheetWidth = int.Parse(SheetDefinition.getProperty("z_sheet_width", "0"));
            // シート高さ取得
            int sheetHeight = int.Parse(SheetDefinition.getProperty("z_sheet_height", "0"));

            // 列幅設定
            for (int i = 1; i <= sheetWidth; i++)
            {
                ExportSheet.Column(i).Width = Worksheets[templateSheetIndex].Column(i).Width;
            }

            // 印字向きをテンプレートシートに合わせる
            ExportSheet.PrinterSettings.Orientation = Worksheets[templateSheetIndex].PrinterSettings.Orientation;
            // 用紙サイズを設定
            ExportSheet.PrinterSettings.PaperSize = Worksheets[templateSheetIndex].PrinterSettings.PaperSize;
            // すべての列を１ページに印刷
            ExportSheet.PrinterSettings.FitToPage = true;
            // 列幅調整あり
            ExportSheet.PrinterSettings.FitToWidth = 1;
            // 行高さ調整なし
            ExportSheet.PrinterSettings.FitToHeight = 0;
            // 余白
            ExportSheet.PrinterSettings.TopMargin = 0;
            ExportSheet.PrinterSettings.BottomMargin = 0;
            ExportSheet.PrinterSettings.LeftMargin = 0;
            ExportSheet.PrinterSettings.RightMargin = 0;
            // 水平垂直
            ExportSheet.PrinterSettings.HorizontalCentered = true;
            ExportSheet.PrinterSettings.VerticalCentered = true;
            // 枠線の非表示
            ExportSheet.View.ShowGridLines = false;
            // 改ページプレビューなし
            ExportSheet.View.PageBreakView = false;
            // 印刷範囲の設定
            ExportSheet.PrinterSettings.PrintArea = ExportSheet.Cells[1, 1, sheetHeight * (CurrentPageCount), sheetWidth];


        }

        /// <summary>
		/// 各シートインデックスに対応する出力定義取得
		/// </summary>
		/// <param name="outputData">出力データ</outputData></param>
		private Item GetExportDefinitionsAt(string sheetIndex, string blockIndex)
        {
            //return ExportDefinition.getItemsByXPath("//Item[z_sheet_index='" + sheetIndex + "' and z_block_index='" + blockIndex + "']");
            string reportDefinitionID = ReportDefinition.getID();
            Item res = CallProcedure("z_get_export_definition", new string[] { reportDefinitionID, sheetIndex, blockIndex });

            return res;
        }

        /// <summary>
        /// シート定義を取得
        /// </summary>
        /// <param name="outputData">出力データ</outputData></param>
        private Item GetSheetDefinitionsAt(string sheetIndex)
        {
            Item sheetDef = Inn.newItem("z_rel_report_sheet_definition", Actions.GET);
            sheetDef.setProperty("source_id", ReportDefinition.getID());
            sheetDef.setProperty("z_sheet_index", sheetIndex);
            return sheetDef.apply();
        }

        /// <summary>
		/// 操作中セル位置設定
		/// </summary>
		/// <param name="x">列番号</param>
		/// <param name="y">行番号</param>
		/// <returns>改ページ有無</returns>
		private bool SetCurrentPosition(int x, int y)
        {
            if (CurrentPosition == null)
            {
                CurrentPosition = new Point(1, 1);
            }

            CurrentPosition.X = x;
            CurrentPosition.Y = y;

            int startRow = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_DETAIL_START_ROW_EXPORT, "0"));
            int avairableHeight = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_SHEET_HEIGHT, "0"));

            // 1ページ目または最大明細行を出力した場合改ページ
            if (TotalPageCount == 0 || IsBreakRow(CurrentPosition.Y) || CurrentPageCount == 0)
            {
                // シートをコピー
                CopyTemplateToWorkSheet();

                // 各ページの先頭行 + 明細印字開始位置を操作中セルとする
                CurrentPosition.Y = avairableHeight * (CurrentPageCount - 1) + startRow;
                CurrentPosition.X = 1;

                return true;
            }

            return false;
        }

        /// <summary>
		/// 改ページ行判定
		/// </summary>
		/// <param name="currentRow">現在行番号</param>
		/// <returns>改ページ有無</returns>
		private bool IsBreakRow(int currentRow)
        {
            int startRow = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_DETAIL_START_ROW_EXPORT, "0"));
            int maxRowCount = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_MAX_DETAIL_ROWS, "0"));
            int avairableHeight = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_SHEET_HEIGHT, "0"));
            int breakRow = (CurrentPageCount <= 1 ? 0 : CurrentPageCount - 1) * avairableHeight + startRow + maxRowCount;
            return breakRow <= currentRow;
        }

        /// <summary>
		/// テンプレートシートをコピーし次ページを追加。
		/// </summary>
		private void CopyTemplateToWorkSheet()
        {
            int endRow = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_SHEET_HEIGHT, "0"));
            int endColumn = int.Parse(SheetDefinition.getProperty(Items.Report.SheetDef.Props.Z_SHEET_WIDTH, "0"));
            int startRow = endRow * CurrentPageCount + 1;

            int sheetIndex = int.Parse(SheetDefinition.getProperty("z_sheet_index", "1"));

            // 特定範囲をコピー
            ExcelRange sourceRange = Worksheets[sheetIndex].Cells[1, 1, endRow, endColumn];
            sourceRange.Copy(ExportSheet.Cells[startRow, 1]);

            // 行高さを複製
            for (int i = 0; i < endRow; i++)
            {
                ExportSheet.Row(startRow + i).Height = Worksheets[sheetIndex].Row(i + 1).Height;
            }

            // 列幅を複製
            for (int i = 1; i <= endColumn; i++)
            {
                ExportSheet.Column(i).Width = Worksheets[sheetIndex].Column(i).Width;
            }

            // 総ページ数をインクリメント
            TotalPageCount++;
            CurrentPageCount++;

            // 改ページ位置を設定
            ExportSheet.Row(endRow * CurrentPageCount).PageBreak = true;
        }

        /// <summary>
        /// Excel水平配置の取得
        /// </summary>
        /// <param name="def"></param>
        private ExcelHorizontalAlignment GetExcelHorizontalAlignment(string def)
        {
            switch (def)
            {
                case "Center":
                    return ExcelHorizontalAlignment.Center;
                    break;

                case "Right":
                    return ExcelHorizontalAlignment.Right;
                    break;

                case "Left":
                    return ExcelHorizontalAlignment.Left;
                    break;

                default:
                    return ExcelHorizontalAlignment.Left;
                    break;
            }
        }

        /// <summary>
        /// Identityの名称取得
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private string GetIdentityName(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                errMsg = "IdentityのIDが指定されていません。";
                throw new Exception("errMsg");
            }

            Item identityItem = Inn.getItemById("Identity", id);
            string identityName = identityItem.getProperty("name", "");

            return identityName;
        }

        /// <summary>
        /// 作業フォルダの削除
        /// </summary>
        private void DeleteWorkDirectory()
        {
            batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO, "作業フォルダ削除");
            try
            {
                ExportExcel.Dispose();
                Directory.Delete(TempFolderPath, true);
            }
            catch (Exception ex)
            {
                batchLogger.WriteLog(BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.ERROR, "作業フォルダ削除失敗" + ex.ToString());

            }
        }

        /// <summary>
        /// プロシージャの実行
        /// </summary>
        /// <param name="procedureName">プロシージャ名</param>
        private Item CallProcedure(string procedureName, params string[] parameters)
        {
            Item result = Inn.newItem("SQL", "SQL PROCESS");
            result.setProperty("name", procedureName);
            result.setProperty("PROCESS", "CALL");
            if (parameters != null)
            {
                for (int i = 0; i < parameters.Length; i++)
                {
                    result.setProperty("ARG" + (i + 1), parameters[i]);
                }
            }

            return result.apply();

        }
        #region ログ出力関連

        /// <summary>
        /// ユーザメッセージを取得する
        /// </summary>
        /// <param name="userMessageName">ユーザメッセージ</param>
        /// <param name="prms">パラメータ</param>
        /// <returns></returns>
        private string GetUserMessage(string userMessageName, params string[] prms)
        {
            return RGKCustomize.CommonTools.GetUserMessage(Inn, ThisCCO, userMessageName, prms);
        }

        /// <summary>
        /// システムログ出力の初期化処理
        /// </summary>
        /// <param name="logOutputDefinition">z_log_output_definitionで定義したログ出力定義名</param>
        private void InitSystemLogger(string logOutputDefinition)
        {
            if (_logger == null)
            {
                _logger = new BatchLog.Logger(Inn, ThisCCO, logOutputDefinition);
            }
        }

        /// <summary>
        /// システムログ出力の基本処理(★★★直接呼ばないこと★★★)
        /// </summary>
        /// <param name="logMessage"></param>
        /// <param name="logLevel"></param>
        private void SystemLogBase(string logMessage, BatchLog.Logger.LogFileType logType, BatchLog.Logger.LogLevel logLevel)
        {
            // ログ出力
            _logger.WriteLog(logType, logLevel, logMessage);
        }

        /// <summary>
        /// システムインフォメーションログ出力
        /// </summary>
        /// <param name="logMessage"></param>
        private void LogSystemInfo(string logMessage)
        {
            LogDebug(logMessage);
            SystemLogBase(logMessage, BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO);
        }

        /// <summary>
        /// システムエラーログ出力
        /// </summary>
        /// <param name="logMessage"></param>
        private void LogSystemError(string logMessage)
        {
            LogError(logMessage);
            SystemLogBase(logMessage, BatchLog.Logger.LogFileType.SYSTEM_LOG, BatchLog.Logger.LogLevel.INFO);
            SystemLogBase(logMessage, BatchLog.Logger.LogFileType.ERROR_LOG, BatchLog.Logger.LogLevel.ERROR);
        }

        /// <summary>
        /// インフォメーションログ出力
        /// </summary>
        /// <param name="message">メッセージ</param>
        private void LogInfo(string logMessage)
        {
            OutputAppLog("INFO-LOG", logMessage);
        }

        /// <summary>
        /// エラーログ出力
        /// </summary>
        /// <param name="logMessage">メッセージ</param>
        /// <param name="errType">エラーの種類</param>
        private void LogError(string logMessage, string errType = null)
        {
            if (String.IsNullOrWhiteSpace(errType))
            {
                errType = "System Error : ";
            }
            OutputAppLog(errType, logMessage);
        }

        // ---DEBUG--- 
        /// <summary>
        /// デバッグログ出力
        /// </summary>
        /// <param name="logMessage"></param>
        private void LogDebug(string logMessage)
        {
            if (_ENABLE_DEBUG_LOG_)
            {
                OutputAppLog("DEBUG-LOG", logMessage);
            }
        }

        /// <summary>
        /// デバッグログ出力(JSON形式)
        /// </summary>
        /// <param name="logObject"></param>
        private void LogDebugJson(string logMessage, object logObject)
        {
            if (_ENABLE_DEBUG_LOG_)
            {
                // JSON形式でログ出力
                string jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(logObject);
                OutputAppLog("DEBUG-LOG", logMessage + "：" + jsonString);
            }
        }

        /// <summary>
        /// ログ出力(★★★直接呼ばないこと★★★)
        /// </summary>
        /// <param name="logType">ログタイプ</param>
        /// <param name="logString">出力内容</param>
        private void OutputAppLog(string logType, string logString)
        {
            if (!_ENABLE_DEBUG_LOG_)
            {
                RGKCustomize.CommonTools.AppLog(Inn, ThisCCO, METHOD_NAME, logType, logString);
            }
            else
            {
                // デバッグ用ログ出力処理
                //日付ごと
                string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
                // エラーログ用フォーマット
                const string ERROR_LOG = "[{0}][{1}][{2}][{3}]";
                // ログインユーザ名取得用AML
                const string AML_GET_LOGIN_NAME =
                    "<AML>" +
                    "  <Item type='User' action='get' id='{0}' select='login_name,first_name,last_name'></Item>" +
                    "</AML>";
                string loginUserName = "";

                // ログインユーザ名を取得
                Item loginUser = Inn.applyAML(string.Format(AML_GET_LOGIN_NAME, Inn.getUserID()));
                if (!loginUser.isError())
                {
                    loginUserName = loginUser.getProperty("login_name", "");
                    loginUserName += "_" + loginUser.getProperty("last_name", "");
                    loginUserName += loginUser.getProperty("first_name", "");
                }

                // エラーログ用ファイル名
                //const string FILE_NAME = "application_{0}";
                string FILE_NAME = METHOD_NAME + "_{0}_" + loginUserName.Replace("/", "_");
                // 出力内容を作成
                string outputLog = string.Format(ERROR_LOG, METHOD_NAME, loginUserName, logType, logString);
                // エラーログを出力
                ThisCCO.Utilities.WriteDebug(string.Format(FILE_NAME, currentDate), outputLog);
            }
        }

        #endregion ログ出力関連

        private void Dummy()
        {
            // }

