using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office
{
    public class ExcelDocument : IDisposable
    {
        #region 枚举
        #region Excel文件类型
        /// <summary>
        /// Excel文件类型。
        /// </summary>
        public enum ExcelFileFormat
        {
            /// <summary>
            /// Specifies a type of text format。
            /// </summary>
            xlCurrentPlatformText = -4158,
            /// <summary>
            /// Excel workbook format.
            /// </summary>
            xlWorkbookNormal = -4143,
            /// <summary>
            /// Symbolic link format.
            /// </summary>
            xlSYLK = 2,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWKS = 4,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK1 = 5,
            /// <summary>
            /// Comma separated value.
            /// </summary>
            xlCSV = 6,
            /// <summary>
            /// Dbase 2 format.
            /// </summary>
            xlDBF2 = 7,
            /// <summary>
            /// Dbase 3 format.
            /// </summary>
            xlDBF3 = 8,
            /// <summary>
            /// Data Interchange format.
            /// </summary>
            xlDIF = 9,
            /// <summary>
            /// Dbase 4 format.
            /// </summary>
            xlDBF4 = 11,
            /// <summary>
            /// Deprecated format.
            /// </summary>
            xlWJ2WD1 = 14,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK3 = 15,
            /// <summary>
            /// Excel version 2.0.
            /// </summary>
            xlExcel2 = 16,
            /// <summary>
            /// Excel template format.
            /// </summary>
            xlTemplate = 17,
            /// <summary>
            /// Template 8
            /// </summary>
            xlTemplate8 = 17,
            /// <summary>
            /// Microsoft Office Excel Add-In.
            /// </summary>
            xlAddIn = 18,
            /// <summary>
            /// Microsoft Excel 97-2003 Add-In
            /// </summary>
            xlAddIn8 = 18,
            /// <summary>
            /// Specifies a type of text format.
            /// </summary>
            xlTextMac = 19,
            /// <summary>
            /// Specifies a type of text format.
            /// </summary>
            xlTextWindows = 20,
            /// <summary>
            /// Specifies a type of text format.
            /// </summary>
            xlTextMSDOS = 21,
            /// <summary>
            /// Comma separated value.
            /// </summary>
            xlCSVMac = 22,
            /// <summary>
            /// Comma separated value.
            /// </summary>
            xlCSVWindows = 23,
            /// <summary>
            /// Comma separated value.
            /// </summary>
            xlCSVMSDOS = 24,
            /// <summary>
            /// Deprecated format.
            /// </summary>
            xlIntlMacro = 25,
            /// <summary>
            /// Microsoft Office Excel Add-In international format.
            /// </summary>
            xlIntlAddIn = 26,
            /// <summary>
            /// Excel version 2.0 far east.
            /// </summary>
            xlExcel2FarEast = 27,
            /// <summary>
            /// Microsoft Works 2.0 format
            /// </summary>
            xlWorks2FarEast = 28,
            /// <summary>
            /// Excel version 3.0.
            /// </summary>
            xlExcel3 = 29,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK1FMT = 30,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK1ALL = 31,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK3FM3 = 32,
            /// <summary>
            /// Excel version 4.0.
            /// </summary>
            xlExcel4 = 33,
            /// <summary>
            /// Quattro Pro format.
            /// </summary>
            xlWQ1 = 34,
            /// <summary>
            /// Excel version 4.0. Workbook format.
            /// </summary>
            xlExcel4Workbook = 35,
            /// <summary>
            /// Specifies a type of text format.
            /// </summary>
            xlTextPrinter = 36,
            /// <summary>
            /// Lotus 1-2-3 format.
            /// </summary>
            xlWK4 = 38,
            /// <summary>
            /// Excel 95.
            /// </summary>
            xlExcel7 = 39,
            /// <summary>
            /// Excel version 5.0.
            /// </summary>
            xlExcel5 = 39,
            /// <summary>
            /// Deprecated format.
            /// </summary>
            xlWJ3 = 40,
            /// <summary>
            /// Deprecated format.
            /// </summary>
            xlWJ3FJ3 = 41,
            /// <summary>
            /// Specifies a type of text format.
            /// </summary>
            xlUnicodeText = 42,
            /// <summary>
            /// Excel version 95 and 97.
            /// </summary>
            xlExcel9795 = 43,
            /// <summary>
            /// Web page format.
            /// </summary>
            xlHtml = 44,
            /// <summary>
            /// MHT format.
            /// </summary>
            xlWebArchive = 45,
            /// <summary>
            /// Excel Spreadsheet format.
            /// </summary>
            xlXMLSpreadsheet = 46,
            /// <summary>
            /// Excel12
            /// </summary>
            xlExcel12 = 50,
            /// <summary>
            /// Open XML Workbook
            /// </summary>
            xlOpenXMLWorkbook = 51,
            /// <summary>
            /// Workbook default
            /// </summary>
            xlWorkbookDefault = 51,
            /// <summary>
            /// Open XML Workbook Macro Enabled
            /// </summary>
            xlOpenXMLWorkbookMacroEnabled = 52,
            /// <summary>
            /// Open XML Template Macro Enabled
            /// </summary>
            xlOpenXMLTemplateMacroEnabled = 53,
            /// <summary>
            /// Open XML Template
            /// </summary>
            xlOpenXMLTemplate = 54,
            /// <summary>
            /// Open XML Add-In
            /// </summary>
            xlOpenXMLAddIn = 55,
            /// <summary>
            /// Excel8
            /// </summary>
            xlExcel8 = 56,
            /// <summary>
            /// OpenDocument Spreadsheet
            /// </summary>
            xlOpenDocumentSpreadsheet = 60,
        }
        #endregion

        #region 水平对齐方式
        /// <summary>
        /// 水平对齐方式。
        /// </summary>
        public enum ExcelHAlign
        {
            /// <summary>
            /// Right.
            /// <para>靠右（缩进）。</para>
            /// </summary>
            xlHAlignRight = -4152,
            /// <summary>
            /// Left.
            /// <para>靠左（缩进）。</para>
            /// </summary>
            xlHAlignLeft = -4131,
            /// <summary>
            /// Justify.
            /// <para>两端对齐。</para>
            /// </summary>
            xlHAlignJustify = -4130,
            /// <summary>
            /// Distribute.
            /// <para>分散对齐（缩进）。</para>
            /// </summary>
            xlHAlignDistributed = -4117,
            /// <summary>
            /// Center.
            /// <para>居中。</para>
            /// </summary>
            xlHAlignCenter = -4108,
            /// <summary>
            /// Align according to data type.
            /// <para>常规。</para>
            /// </summary>
            xlHAlignGeneral = 1,
            /// <summary>
            /// Fill.
            /// <para>填充。</para>
            /// </summary>
            xlHAlignFill = 5,
            /// <summary>
            /// Center across selection.
            /// <para>跨列居中。</para>
            /// </summary>
            xlHAlignCenterAcrossSelection = 7,
        }
        #endregion

        #region 垂直对齐方式
        /// <summary>
        /// 垂直对齐方式。
        /// </summary>
        public enum ExcelVAlign
        {
            /// <summary>
            /// Top.
            /// <para>靠上。</para>
            /// </summary>
            xlVAlignTop = -4160,
            /// <summary>
            /// Justify.
            /// <para>两端对齐。</para>
            /// </summary>
            xlVAlignJustify = -4130,
            /// <summary>
            /// Distributed.
            /// <para>分散对齐。</para>
            /// </summary>
            xlVAlignDistributed = -4117,
            /// <summary>
            /// Center.
            /// <para>居中。</para>
            /// </summary>
            xlVAlignCenter = -4108,
            /// <summary>
            /// Bottom.
            /// <para>靠下。</para>
            /// </summary>
            xlVAlignBottom = -4107,
        }
        #endregion

        #region 颜色类型
        /// <summary>
        /// 颜色类型。
        /// </summary>
        public enum ExcelColorIndex
        {
            /// <summary>
            /// No color（无颜色）。
            /// </summary>
            xlColorIndexNone = -4142,
            /// <summary>
            /// Automatic color（自动）。
            /// </summary>
            xlColorIndexAutomatic = -4105,
            /// <summary>
            /// Black（黑色）。
            /// </summary>
            xlBlack = 1,
            /// <summary>
            /// DarkRed（深红）。
            /// </summary>
            xlDarkRed = 9,
            /// <summary>
            /// Red（红色）。
            /// </summary>
            xlRed = 3,
            /// <summary>
            /// Pink（粉红）。
            /// </summary>
            xlPink = 7,
            /// <summary>
            /// Rose（玫瑰红）。
            /// </summary>
            xlRose = 38,
            /// <summary>
            /// Brown（棕色）。
            /// </summary>
            xlBrown = 53,
            /// <summary>
            /// Orange（橙色）。
            /// </summary>
            xlOrange = 46,
            /// <summary>
            /// LightOrange（浅橙色）。
            /// </summary>
            xlLightOrange = 45,
            /// <summary>
            /// Gold（金色）。
            /// </summary>
            xlGold = 44,
            /// <summary>
            /// DarkBrown（茶色）。
            /// </summary>
            xlDarkBrown = 40,
            /// <summary>
            /// Olive（橄榄色）。
            /// </summary>
            xlOlive = 52,
            /// <summary>
            /// DarkYellow（深黄）。
            /// </summary>
            xlDarkYellow = 12,
            /// <summary>
            /// Lime（石灰色）。
            /// </summary>
            xlLime = 43,
            /// <summary>
            /// Yellow（黄色）。
            /// </summary>
            xlYellow = 6,
            /// <summary>
            /// LightYellow（浅黄色）。
            /// </summary>
            xlLightYellow = 36,
            /// <summary>
            /// DarkGreen（深绿色）。
            /// </summary>
            xlDarkGreen = 51,
            /// <summary>
            /// Green（绿色）。
            /// </summary>
            xlGreen = 10,
            /// <summary>
            /// SeaGreen（海绿色）。
            /// </summary>
            xlSeaGreen = 50,
            /// <summary>
            /// BrightGreen（荧光绿）。
            /// </summary>
            xlBrightGreen = 4,
            /// <summary>
            /// LightGreen（浅绿色）。
            /// </summary>
            xlLightGreen = 35,
            /// <summary>
            /// DarkCyan（暗青色）。
            /// </summary>
            xlDarkCyan = 49,
            /// <summary>
            /// Cyan（青色）。
            /// </summary>
            xlCyan = 14,
            /// <summary>
            /// WaterGreen（水绿色）。
            /// </summary>
            xlWaterGreen = 42,
            /// <summary>
            /// CyanGreen（青绿色）。
            /// </summary>
            xlCyanGreen = 8,
            /// <summary>
            /// LightCyanGreen（浅青绿色）。
            /// </summary>
            xlLightCyanGreen = 34,
            /// <summary>
            /// DarkBlue（深蓝）。
            /// </summary>
            xlDarkBlue = 11,
            /// <summary>
            /// Blue（蓝色）。
            /// </summary>
            xlBlue = 5,
            /// <summary>
            /// LightSkyBlue（浅天蓝色）。
            /// </summary>
            xlLightSkyBlue = 41,
            /// <summary>
            /// SkyBlue（天蓝色）。
            /// </summary>
            xlSkyBlue = 33,
            /// <summary>
            /// LightBlue（浅蓝色）。
            /// </summary>
            xlLightBlue = 37,
            /// <summary>
            /// Indigo（靛青色）。
            /// </summary>
            xlIndigo = 55,
            /// <summary>
            /// BlueGray（蓝灰色）。
            /// </summary>
            xlBlueGray = 47,
            /// <summary>
            /// Violet（紫色）。
            /// </summary>
            xlViolet = 13,
            /// <summary>
            /// Plum（梅红）。
            /// </summary>
            xlPlum = 54,
            /// <summary>
            /// LightViolet（浅紫色）。
            /// </summary>
            xlLightViolet = 39,
            /// <summary>
            /// Gray80（80%灰色）。
            /// </summary>
            xlGray80 = 56,
            /// <summary>
            /// Gray50（50%灰色）。
            /// </summary>
            xlGray50 = 16,
            /// <summary>
            /// Gray40（40%灰色）。
            /// </summary>
            xlGray40 = 48,
            /// <summary>
            /// Gray25（25%灰色）。
            /// </summary>
            xlGray25 = 15,
            /// <summary>
            /// White（白色）。
            /// </summary>
            xlWhite = 2
        }
        #endregion

        #region 线条样式
        /// <summary>
        /// 线条样式。
        /// </summary>
        public enum ExcelLineStyle
        {
            /// <summary>
            /// No line（无线）。
            /// </summary>
            xlLineStyleNone = -4142,
            /// <summary>
            /// Double line（双线）。
            /// </summary>
            xlDouble = -4119,
            /// <summary>
            /// Dotted line（虚线）。
            /// </summary>
            xlDot = -4118,
            /// <summary>
            /// Dashed line（虚线）。
            /// </summary>
            xlDash = -4115,
            /// <summary>
            /// Continuous line（连续线）。
            /// </summary>
            xlContinuous = 1,
            /// <summary>
            /// Alternating dashes and dots（单点和虚线交替）。
            /// </summary>
            xlDashDot = 4,
            /// <summary>
            /// Dash followed by two dots（双点和虚线交替）。
            /// </summary>
            xlDashDotDot = 5,
            /// <summary>
            /// Slanted dashes（倾斜的破折号）。
            /// </summary>
            xlSlantDashDot = 13,
        }
        #endregion

        #region 边框粗细
        /// <summary>
        /// 边框粗细。
        /// </summary>
        public enum ExcelBorderWeight
        {
            /// <summary>
            /// Medium（中等）。
            /// </summary>
            xlMedium = -4138,
            /// <summary>
            /// Hairline (thinnest border)（极细）。
            /// </summary>
            xlHairline = 1,
            /// <summary>
            /// Thin（细）。
            /// </summary>
            xlThin = 2,
            /// <summary>
            /// Thick (widest border)（粗）。
            /// </summary>
            xlThick = 4,
        }
        #endregion

        #region 边框样式
        /// <summary>
        /// 边框样式。
        /// </summary>
        public enum ExcelBordersIndex
        {
            /// <summary>
            /// Border running from the upper left-hand corner to the lower right of each cell in the range.
            /// <para>从左上到右下的斜线。</para>
            /// </summary>
            xlDiagonalDown = 5,
            /// <summary>
            /// Border running from the lower left-hand corner to the upper right of each cell in the range.
            /// <para>从左下到右上的斜线。</para>
            /// </summary>
            xlDiagonalUp = 6,
            /// <summary>
            /// Border at the left-hand edge of the range.
            /// <para>左边框。</para>
            /// </summary>
            xlEdgeLeft = 7,
            /// <summary>
            /// Border at the top of the range.
            /// <para>上边框。</para>
            /// </summary>
            xlEdgeTop = 8,
            /// <summary>
            /// Border at the bottom of the range.
            /// <para>底边框。</para>
            /// </summary>
            xlEdgeBottom = 9,
            /// <summary>
            /// Border at the right-hand edge of the range.
            /// <para>右边框。</para>
            /// </summary>
            xlEdgeRight = 10,
            /// <summary>
            /// Vertical borders for all the cells in the range except borders on the outside of the range.
            /// <para>水平边框（上边框+下边框）。</para>
            /// </summary>
            xlInsideVertical = 11,
            /// <summary>
            /// Horizontal borders for all cells in the range except borders on the outside of the range.
            /// <para>垂直边框（左边框+右边框）</para>
            /// </summary>
            xlInsideHorizontal = 12,
        }
        #endregion

        #region 底纹样式
        /// <summary>
        /// 底纹样式。
        /// </summary>
        public enum ExcelPattern
        {
            /// <summary>
            /// Dark vertical bars.
            /// <para>垂直条纹。</para>
            /// </summary>
            xlPatternVertical = -4166,
            /// <summary>
            /// Dark diagonal lines running from the lower left to the upper right.
            /// <para>对角线条纹。</para>
            /// </summary>
            xlPatternUp = -4162,
            /// <summary>
            /// No pattern.
            /// <para>无底纹。</para>
            /// </summary>
            xlPatternNone = -4142,
            /// <summary>
            /// Dark horizontal lines.
            /// <para>水平条纹。</para>
            /// </summary>
            xlPatternHorizontal = -4128,
            /// <summary>
            /// 75% gray.
            /// <para>75% 灰色。</para>
            /// </summary>
            xlPatternGray75 = -4126,
            /// <summary>
            /// 50% gray.
            /// <para>50% 灰色。</para>
            /// </summary>
            xlPatternGray50 = -4125,
            /// <summary>
            /// 25% gray.
            /// <para>25% 灰色。</para>
            /// </summary>
            xlPatternGray25 = -4124,
            /// <summary>
            /// Dark diagonal lines running from the upper left to the lower right.
            /// <para>逆对角线条纹。</para>
            /// </summary>
            xlPatternDown = -4121,
            /// <summary>
            /// Excel controls the pattern.
            /// <para>自动。</para>
            /// </summary>
            xlPatternAutomatic = -4105,
            /// <summary>
            /// Solid color.
            /// <para>实心。</para>
            /// </summary>
            xlPatternSolid = 1,
            /// <summary>
            /// Checkerboard.
            /// <para>对角线刨面线。</para>
            /// </summary>
            xlPatternChecker = 9,
            /// <summary>
            /// 75% dark moiré.
            /// <para>粗对角线剖面线。</para>
            /// </summary>
            xlPatternSemiGray75 = 10,
            /// <summary>
            /// Light horizontal lines.
            /// <para>细水平条纹。</para>
            /// </summary>
            xlPatternLightHorizontal = 11,
            /// <summary>
            /// Light vertical bars.
            /// <para>细垂直条纹。</para>
            /// </summary>
            xlPatternLightVertical = 12,
            /// <summary>
            /// Light diagonal lines running from the upper left to the lower right.
            /// <para>细逆对角线条纹。</para>
            /// </summary>
            xlPatternLightDown = 13,
            /// <summary>
            /// Light diagonal lines running from the lower left to the upper right.
            /// <para>细对角线条纹。</para>
            /// </summary>
            xlPatternLightUp = 14,
            /// <summary>
            /// Grid.
            /// <para>细水平剖面线。</para>
            /// </summary>
            xlPatternGrid = 15,
            /// <summary>
            /// Criss-cross lines.
            /// <para>细对角线剖面线。</para>
            /// </summary>
            xlPatternCrissCross = 16,
            /// <summary>
            /// 16% gray.
            /// <para>12.5% 灰色。</para>
            /// </summary>
            xlPatternGray16 = 17,
            /// <summary>
            /// 8% gray.
            /// <para>6.25% 灰色。</para>
            /// </summary>
            xlPatternGray8 = 18,
            /// <summary>
            /// 线性渐变。
            /// </summary>
            xlPatternLinearGradient = 4000,
            /// <summary>
            /// 矩形梯度渐变。
            /// </summary>
            xlPatternRectangularGradient = 4001,
        }
        #endregion

        #region 复制粘贴类型
        /// <summary>
        /// 复制粘贴类型。
        /// </summary>
        public enum ExcelPasteType
        {
            /// <summary>
            /// 值。
            /// </summary>
            xlPasteValues = -4163,
            /// <summary>
            /// 批注。
            /// </summary>
            xlPasteComments = -4144,
            /// <summary>
            /// 公式。
            /// </summary>
            xlPasteFormulas = -4123,
            /// <summary>
            /// 格式。
            /// </summary>
            xlPasteFormats = -4122,
            /// <summary>
            /// 全部。
            /// </summary>
            xlPasteAll = -4104,
            /// <summary>
            /// 有效性验证。
            /// </summary>
            xlPasteValidation = 6,
            /// <summary>
            /// 边框除外。
            /// </summary>
            xlPasteAllExceptBorders = 7,
            /// <summary>
            /// 列宽。
            /// </summary>
            xlPasteColumnWidths = 8,
            /// <summary>
            /// 公式和数字格式。
            /// </summary>
            xlPasteFormulasAndNumberFormats = 11,
            /// <summary>
            /// 值和数字格式。
            /// </summary>
            xlPasteValuesAndNumberFormats = 12,
            /// <summary>
            /// 所有使用源主题的单元。
            /// </summary>
            xlPasteAllUsingSourceTheme = 13,
            /// <summary>
            /// 所有内容并且将合并条件格式。
            /// </summary>
            xlPasteAllMergingConditionalFormats = 14,
        }
        #endregion

        #region 纸张类型
        /// <summary>
        /// 纸张类型。
        /// </summary>
        public enum ExcelPaperSize
        {
            /// <summary>
            /// Letter (8-1/2 in. x 11 in.)
            /// </summary>
            xlPaperLetter = 1,
            /// <summary>
            /// Letter Small (8-1/2 in. x 11 in.)
            /// </summary>
            xlPaperLetterSmall = 2,
            /// <summary>
            /// Tabloid (11 in. x 17 in.)
            /// </summary>
            xlPaperTabloid = 3,
            /// <summary>
            /// Ledger (17 in. x 11 in.)
            /// </summary>
            xlPaperLedger = 4,
            /// <summary>
            /// Legal (8-1/2 in. x 14 in.)
            /// </summary>
            xlPaperLegal = 5,
            /// <summary>
            /// Statement (5-1/2 in. x 8-1/2 in.)
            /// </summary>
            xlPaperStatement = 6,
            /// <summary>
            /// Executive (7-1/2 in. x 10-1/2 in.)
            /// </summary>
            xlPaperExecutive = 7,
            /// <summary>
            /// A3 (297 mm x 420 mm)
            /// </summary>
            xlPaperA3 = 8,
            /// <summary>
            /// A4 (210 mm x 297 mm)
            /// </summary>
            xlPaperA4 = 9,
            /// <summary>
            /// A4 Small (210 mm x 297 mm)
            /// </summary>
            xlPaperA4Small = 10,
            /// <summary>
            /// A5 (148 mm x 210 mm)
            /// </summary>
            xlPaperA5 = 11,
            /// <summary>
            /// B4 (250 mm x 354 mm)
            /// </summary>
            xlPaperB4 = 12,
            /// <summary>
            /// A5 (148 mm x 210 mm)
            /// </summary>
            xlPaperB5 = 13,
            /// <summary>
            /// Folio (8-1/2 in. x 13 in.)
            /// </summary>
            xlPaperFolio = 14,
            /// <summary>
            /// Quarto (215 mm x 275 mm)
            /// </summary>
            xlPaperQuarto = 15,
            /// <summary>
            /// 10 in. x 14 in.
            /// </summary>
            xlPaper10x14 = 16,
            /// <summary>
            /// 11 in. x 17 in.
            /// </summary>
            xlPaper11x17 = 17,
            /// <summary>
            /// Note (8-1/2 in. x 11 in.)
            /// </summary>
            xlPaperNote = 18,
            /// <summary>
            /// Envelope #9 (3-7/8 in. x 8-7/8 in.)
            /// </summary>
            xlPaperEnvelope9 = 19,
            /// <summary>
            /// Envelope #10 (4-1/8 in. x 9-1/2 in.)
            /// </summary>
            xlPaperEnvelope10 = 20,
            /// <summary>
            /// Envelope #11 (4-1/2 in. x 10-3/8 in.)
            /// </summary>
            xlPaperEnvelope11 = 21,
            /// <summary>
            /// Envelope #12 (4-1/2 in. x 11 in.)
            /// </summary>
            xlPaperEnvelope12 = 22,
            /// <summary>
            /// Envelope #14 (5 in. x 11-1/2 in.)
            /// </summary>
            xlPaperEnvelope14 = 23,
            /// <summary>
            /// C size sheet
            /// </summary>
            xlPaperCsheet = 24,
            /// <summary>
            /// D size sheet
            /// </summary>
            xlPaperDsheet = 25,
            /// <summary>
            /// E size sheet
            /// </summary>
            xlPaperEsheet = 26,
            /// <summary>
            /// Envelope DL (110 mm x 220 mm)
            /// </summary>
            xlPaperEnvelopeDL = 27,
            /// <summary>
            /// Envelope C5 (162 mm x 229 mm)
            /// </summary>
            xlPaperEnvelopeC5 = 28,
            /// <summary>
            /// Envelope C3 (324 mm x 458 mm)
            /// </summary>
            xlPaperEnvelopeC3 = 29,
            /// <summary>
            /// Envelope C4 (229 mm x 324 mm)
            /// </summary>
            xlPaperEnvelopeC4 = 30,
            /// <summary>
            /// Envelope C6 (114 mm x 162 mm)
            /// </summary>
            xlPaperEnvelopeC6 = 31,
            /// <summary>
            /// Envelope C65 (114 mm x 229 mm)
            /// </summary>
            xlPaperEnvelopeC65 = 32,
            /// <summary>
            /// Envelope B4 (250 mm x 353 mm)
            /// </summary>
            xlPaperEnvelopeB4 = 33,
            /// <summary>
            /// Envelope B5 (176 mm x 250 mm)
            /// </summary>
            xlPaperEnvelopeB5 = 34,
            /// <summary>
            /// Envelope B6 (176 mm x 125 mm)
            /// </summary>
            xlPaperEnvelopeB6 = 35,
            /// <summary>
            /// Envelope (110 mm x 230 mm)
            /// </summary>
            xlPaperEnvelopeItaly = 36,
            /// <summary>
            /// Envelope Monarch (3-7/8 in. x 7-1/2 in.)
            /// </summary>
            xlPaperEnvelopeMonarch = 37,
            /// <summary>
            /// Envelope (3-5/8 in. x 6-1/2 in.)
            /// </summary>
            xlPaperEnvelopePersonal = 38,
            /// <summary>
            /// U.S. Standard Fanfold (14-7/8 in. x 11 in.)
            /// </summary>
            xlPaperFanfoldUS = 39,
            /// <summary>
            /// German Legal Fanfold (8-1/2 in. x 13 in.)
            /// </summary>
            xlPaperFanfoldStdGerman = 40,
            /// <summary>
            /// German Legal Fanfold (8-1/2 in. x 13 in.)
            /// </summary>
            xlPaperFanfoldLegalGerman = 41,
            /// <summary>
            /// User-defined
            /// </summary>
            xlPaperUser = 256,
        }
        #endregion

        #region 打印方向
        /// <summary>
        /// 打印方向。
        /// </summary>
        public enum ExcelPageOrientation
        {
            /// <summary>
            /// Portrait mode.
            /// <para>纵向。</para>
            /// </summary>
            xlPortrait = 1,
            /// <summary>
            /// Landscape mode.
            /// <para>横向。</para>
            /// </summary>
            xlLandscape = 2,
        }
        #endregion
        #endregion

        #region API函数
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        #endregion

        #region 私有成员

        private Microsoft.Office.Interop.Excel.Application _Excel;

        private bool _IsDisposed = false;

        private int _ProcessID = 0;

        #endregion

        #region 实现接口
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_IsDisposed)// 如果资源未释放 这个判断主要用了防止对象被多次释放
            {
                if (disposing)
                {
                    // 释放托管资源
                }
                // 释放非托管资源
                _Excel.Application.DisplayAlerts = false;
                _Excel.Workbooks.Close();
                _Excel.Application.DisplayAlerts = true;
                _Excel.Quit();
                if (_ProcessID != 0)
                { System.Diagnostics.Process.GetProcessById(_ProcessID).Kill(); }
            }
            _IsDisposed = true; // 标识此对象已释放
        }
        #endregion

        #region 公有属性

        #region 获取或设置Excel对象中当前Sheet的名称
        public string ActiveSheetName
        {
            get { return ((Microsoft.Office.Interop.Excel.Worksheet)_Excel.ActiveSheet).Name; }
            set { ((Microsoft.Office.Interop.Excel.Worksheet)_Excel.ActiveSheet).Name = value; }
        }
        #endregion

        #region 获取对象是否已被释放
        public bool IsDisposed
        {
            get { return _IsDisposed; }
        }
        #endregion

        #region 获取句柄
        public IntPtr Hander
        {
            get { return new IntPtr(_Excel.Hwnd); }
        }
        #endregion

        #region 获取进程ID
        public int ProcessID
        {
            get { return _ProcessID; }
        }
        #endregion

        #endregion

        #region 构造函数

        #region 初始化此类的新实例
        public ExcelDocument()
        {
            _Excel = new Microsoft.Office.Interop.Excel.Application();
            _Excel.Workbooks.Add(Type.Missing);
            GetWindowThreadProcessId(new IntPtr(_Excel.Hwnd), out _ProcessID);
        }

        public ExcelDocument(string templateFileName)
        {
            _Excel = new Microsoft.Office.Interop.Excel.Application();
            _Excel.Workbooks.Add(new System.IO.FileInfo(templateFileName).FullName);
            GetWindowThreadProcessId(new IntPtr(_Excel.Hwnd), out _ProcessID);
        }

        ~ExcelDocument()
        {
            Dispose(false);
        }
        #endregion

        #endregion

        #region 私有方法

        #region 获取单元格
        private Microsoft.Office.Interop.Excel.Range GetRange(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, int sheetIndex = 0)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = sheetIndex == 0 ? _Excel.ActiveSheet : _Excel.Worksheets[sheetIndex];
            Microsoft.Office.Interop.Excel.Range startRange = sheet.Cells[startRowIndex, startColumnIndex];
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells[endRowIndex, endColumnIndex];
            return sheet.get_Range(startRange, endRange);
        }
        #endregion

        #endregion

        #region 共有方法

        #region 显示Excel
        public void SetVisible(bool visible)
        {
            _Excel.Visible = visible;
        }
        #endregion

        #region 关闭Excel
        public void Close()
        {
            Dispose();
        }
        #endregion

        #region 保存文件
        public void SaveAs(string fileName, ExcelFileFormat format = ExcelFileFormat.xlWorkbookDefault)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _Excel.Application.DisplayAlerts = false;
            _Excel.ActiveWorkbook.SaveAs(fullName, (Microsoft.Office.Interop.Excel.XlFileFormat)format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Excel.Application.DisplayAlerts = true;
        }
        #endregion

        #region 保存Excel文件（PDF格式）
        public void SaveAsPDF(string fileName, bool allSheet = false)
        {
            bool openAfterPublish = false;
            string fullName = new System.IO.FileInfo(fileName).FullName;
            if (allSheet)
            {
                _Excel.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, openAfterPublish, Type.Missing);
            }
            else
            {
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
                activeSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, openAfterPublish, Type.Missing);
            }
        }
        #endregion

        #region 预览Excel文件
        public void Preview(bool allSheet = false, bool enableChanges = true)
        {
            if (allSheet)
            {
                _Excel.Worksheets.Select();
                _Excel.Visible = true;
                _Excel.ActiveWindow.SelectedSheets.PrintPreview(enableChanges);
            }
            else
            {
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
                _Excel.Visible = true;
                activeSheet.PrintPreview(enableChanges);
            }
        }
        #endregion

        #region 打印Excel文件
        public void Print(string printerName, bool allSheet = false, int count = 1)
        {
            if (allSheet)
            {
                Microsoft.Office.Interop.Excel.Workbook activeWorkBook = _Excel.ActiveWorkbook;
                activeWorkBook.PrintOutEx(Type.Missing, Type.Missing, count, Type.Missing, printerName, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            else
            {
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
                activeSheet.PrintOut(Type.Missing, Type.Missing, count, Type.Missing, printerName, Type.Missing, Type.Missing, Type.Missing);
            }
        }
        #endregion

        #region 获取Excel对象中的所有Sheet名称
        public string[] GetSheetNames()
        {
            if (_Excel != null)
            {
                return _Excel.Worksheets.Cast<Microsoft.Office.Interop.Excel.Worksheet>().Select<Microsoft.Office.Interop.Excel.Worksheet, string>(sheet => sheet.Name).ToArray();
            }
            else
            {
                return new string[0];
            }
        }
        #endregion

        #region 设置活动Sheet
        public void SetActiveSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Select();
        }
        #endregion

        #region 设置活动Sheet
        public void SetActiveSheet(string sheetName)
        {
            int sheetIndex = GetSheetIndex(sheetName);
            SetActiveSheet(sheetIndex);
        }
        #endregion

        #region 获取指定索引的Sheet名称
        public string GetSheetName(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            return sheet.Name;
        }
        #endregion

        #region 设置指定索引的Sheet名称
        public void SetSheetName(int index, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Name = sheetName;
        }
        #endregion

        #region 获取Sheet数量
        public int GetSheetCount()
        {
            return _Excel.Worksheets.Count;
        }
        #endregion

        #region 获取Sheet的索引
        public int GetSheetIndex(string sheetName)
        {
            return GetSheetNames().ToList().FindIndex(x => x.ToLower() == sheetName.ToLower()) + 1;
        }
        #endregion

        #region 添加一个Sheet
        public void AddSheet()
        {
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            _Excel.Worksheets.Add(Type.Missing, lastSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
        }
        #endregion

        #region 添加一个Sheet
        public void AddSheet(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            _Excel.Worksheets.Add(Type.Missing, lastSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            ActiveSheetName = sheetName;
        }
        #endregion

        #region 添加一个复制的Sheet
        public void AddCopySheet(int originalSheetIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            originalSheet.Copy(Type.Missing, lastSheet);
        }
        #endregion

        #region 添加一个复制的Sheet
        public void AddCopySheet(int originalSheetIndex, string newSheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            originalSheet.Copy(Type.Missing, lastSheet);
            ActiveSheetName = newSheetName;
        }
        #endregion

        #region 插入一个Sheet
        public void InsertSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            this._Excel.Worksheets.Add(sheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
        }
        #endregion

        #region 插入一个Sheet
        public void InsertSheet(int index, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            this._Excel.Worksheets.Add(sheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            ActiveSheetName = sheetName;
        }
        #endregion

        #region 插入一个复制Sheet
        public void InsertCopySheet(int originalSheetIndex, int destinationSheetIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationSheetIndex];
            originalSheet.Copy(destinationSheet, Type.Missing);
        }
        #endregion

        #region 插入一个复制Sheet
        public void InsertCopySheet(int originalSheetIndex, int destinationSheetIndex, string newSheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationSheetIndex];
            originalSheet.Copy(destinationSheet, Type.Missing);
            ActiveSheetName = newSheetName;
        }
        #endregion

        #region 移动Sheet
        public void MoveSheet(int originalIndex, int destinationIndex)
        {
            if (originalIndex == destinationIndex)
            { return; }
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationIndex];
            if (originalIndex < destinationIndex)
            {
                destinationSheet = _Excel.Worksheets[destinationIndex + 1];
            }
            originalSheet.Move(destinationSheet, Type.Missing);
        }
        #endregion

        #region 删除Sheet
        public void DeleteSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Delete();
        }
        #endregion

        #region 删除Sheet
        public void DeleteSheet(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[GetSheetIndex(sheetName)];
            sheet.Delete();
        }
        #endregion

        #region 获取Sheet中已使用的行数
        public int GetUsedRowsCount()
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            return sheet.UsedRange.Rows.Count;
        }
        #endregion

        #region 获取Sheet中已使用的列数
        public int GetUsedColumnCount(int sheetIndex = 0)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            return sheet.UsedRange.Columns.Count;
        }
        #endregion

        #region 设置活动单元格
        public void SetActiveCell(int rowIndex, int columnIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = sheet.Cells[rowIndex, columnIndex];
            range.Select();
            return;
        }
        #endregion

        #region 获取单元格数据
        public object GetCellValue(int rowIndex, int columnIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = activeSheet.Cells[rowIndex, columnIndex];
            return range.Value;
        }
        #endregion

        #region 设置单元格数据
        public void SetCellValue(int rowIndex, int columnIndex, object value)
        {
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = activeSheet.Cells[rowIndex, columnIndex];
            range.Value = value;
        }
        #endregion


        #endregion

    }
}
