using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public struct ExcelColor
    {
        #region 公有属性

        #region 获取颜色的索引
        public int Index { get; private set; }
        #endregion

        #region 获取颜色的NPOI索引
        public short IndexNPOI { get; private set; }
        #endregion

        #region 获取颜色名称
        public string Name { get; private set; }
        #endregion

        #region 获取颜色的描述
        public string Description { get; private set; }
        #endregion

        #region 获取颜色的十六进制代码
        public string HexString { get; private set; }
        #endregion

        #region 获取颜色的alpha分量值
        public int A { get; private set; }
        #endregion

        #region 获取颜色的红色分量值
        public int R { get; private set; }
        #endregion

        #region 获取颜色的绿色分量值
        public int G { get; private set; }
        #endregion

        #region 获取颜色的蓝色分量值
        public int B { get; private set; }
        #endregion

        #region 获取是否是深色系
        public bool IsDarkColor { get; private set; }
        #endregion

        #region 获取已知的颜色
        public static List<ExcelColor> KnownColors { get => GetKnownColor(); }
        #endregion

        #endregion

        #region 静态颜色
        public static ExcelColor None { get; private set; } = new ExcelColor(false);
        public static ExcelColor Black { get; private set; } = new ExcelColor(KnownColor.Black, true);
        public static ExcelColor White { get; private set; } = new ExcelColor(KnownColor.White, false);
        public static ExcelColor Red { get; private set; } = new ExcelColor(KnownColor.Red, true);
        public static ExcelColor BrightGreen { get; private set; } = new ExcelColor(KnownColor.BrightGreen, false);
        public static ExcelColor Blue { get; private set; } = new ExcelColor(KnownColor.Blue, true);
        public static ExcelColor Yellow { get; private set; } = new ExcelColor(KnownColor.Yellow, false);
        public static ExcelColor Pink { get; private set; } = new ExcelColor(KnownColor.Pink, false);
        public static ExcelColor Turquoise { get; private set; } = new ExcelColor(KnownColor.Turquoise, false);
        public static ExcelColor DarkRed { get; private set; } = new ExcelColor(KnownColor.DarkRed, true);
        public static ExcelColor Green { get; private set; } = new ExcelColor(KnownColor.Green, true);
        public static ExcelColor DarkBlue { get; private set; } = new ExcelColor(KnownColor.DarkBlue, true);
        public static ExcelColor DarkYellow { get; private set; } = new ExcelColor(KnownColor.DarkYellow, true);
        public static ExcelColor Violet { get; private set; } = new ExcelColor(KnownColor.Violet, true);
        public static ExcelColor Teal { get; private set; } = new ExcelColor(KnownColor.Teal, true);
        public static ExcelColor Gray25Percent { get; private set; } = new ExcelColor(KnownColor.Gray25Percent, false);
        public static ExcelColor Gray50Percent { get; private set; } = new ExcelColor(KnownColor.Gray50Percent, true);
        public static ExcelColor Periwinkle { get; private set; } = new ExcelColor(KnownColor.Periwinkle, false);
        public static ExcelColor PlumPlus { get; private set; } = new ExcelColor(KnownColor.PlumPlus, true);
        public static ExcelColor Ivory { get; private set; } = new ExcelColor(KnownColor.Ivory, false);
        public static ExcelColor LiteTurquoise { get; private set; } = new ExcelColor(KnownColor.LiteTurquoise, false);
        public static ExcelColor DarkPurple { get; private set; } = new ExcelColor(KnownColor.DarkPurple, true);
        public static ExcelColor Coral { get; private set; } = new ExcelColor(KnownColor.Coral, false);
        public static ExcelColor OceanBlue { get; private set; } = new ExcelColor(KnownColor.OceanBlue, true);
        public static ExcelColor IceBlue { get; private set; } = new ExcelColor(KnownColor.IceBlue, false);
        public static ExcelColor DarkBluePlus { get; private set; } = new ExcelColor(KnownColor.DarkBluePlus, true);
        public static ExcelColor PinkPlus { get; private set; } = new ExcelColor(KnownColor.PinkPlus, false);
        public static ExcelColor YellowPlus { get; private set; } = new ExcelColor(KnownColor.YellowPlus, false);
        public static ExcelColor TurquoisePlus { get; private set; } = new ExcelColor(KnownColor.TurquoisePlus, false);
        public static ExcelColor VioletPlus { get; private set; } = new ExcelColor(KnownColor.VioletPlus, true);
        public static ExcelColor DarkRedPlus { get; private set; } = new ExcelColor(KnownColor.DarkRedPlus, true);
        public static ExcelColor TealPlus { get; private set; } = new ExcelColor(KnownColor.TealPlus, true);
        public static ExcelColor BluePlus { get; private set; } = new ExcelColor(KnownColor.BluePlus, true);
        public static ExcelColor SkyBlue { get; private set; } = new ExcelColor(KnownColor.SkyBlue, false);
        public static ExcelColor LightTurquoise { get; private set; } = new ExcelColor(KnownColor.LightTurquoise, false);
        public static ExcelColor LightGreen { get; private set; } = new ExcelColor(KnownColor.LightGreen, false);
        public static ExcelColor LightYellow { get; private set; } = new ExcelColor(KnownColor.LightYellow, false);
        public static ExcelColor PaleBlue { get; private set; } = new ExcelColor(KnownColor.PaleBlue, false);
        public static ExcelColor Rose { get; private set; } = new ExcelColor(KnownColor.Rose, false);
        public static ExcelColor Lavender { get; private set; } = new ExcelColor(KnownColor.Lavender, false);
        public static ExcelColor Tan { get; private set; } = new ExcelColor(KnownColor.Tan, false);
        public static ExcelColor LightBlue { get; private set; } = new ExcelColor(KnownColor.LightBlue, true);
        public static ExcelColor Aqua { get; private set; } = new ExcelColor(KnownColor.Aqua, false);
        public static ExcelColor Lime { get; private set; } = new ExcelColor(KnownColor.Lime, false);
        public static ExcelColor Gold { get; private set; } = new ExcelColor(KnownColor.Gold, false);
        public static ExcelColor LightOrange { get; private set; } = new ExcelColor(KnownColor.LightOrange, false);
        public static ExcelColor Orange { get; private set; } = new ExcelColor(KnownColor.Orange, true);
        public static ExcelColor BlueGray { get; private set; } = new ExcelColor(KnownColor.BlueGray, true);
        public static ExcelColor Gray40Percent { get; private set; } = new ExcelColor(KnownColor.Gray40Percent, true);
        public static ExcelColor DarkTeal { get; private set; } = new ExcelColor(KnownColor.DarkTeal, true);
        public static ExcelColor SeaGreen { get; private set; } = new ExcelColor(KnownColor.SeaGreen, true);
        public static ExcelColor DarkGreen { get; private set; } = new ExcelColor(KnownColor.DarkGreen, true);
        public static ExcelColor OliveGreen { get; private set; } = new ExcelColor(KnownColor.OliveGreen, true);
        public static ExcelColor Brown { get; private set; } = new ExcelColor(KnownColor.Brown, true);
        public static ExcelColor Plum { get; private set; } = new ExcelColor(KnownColor.Plum, true);
        public static ExcelColor Indigo { get; private set; } = new ExcelColor(KnownColor.Indigo, true);
        public static ExcelColor Gray80Percent { get; private set; } = new ExcelColor(KnownColor.Gray80Percent, true);
        #endregion

        #region 构造函数
        internal ExcelColor(bool isDarkColor)
        {
            Index = 0;
            IndexNPOI = 0;
            Name = "";
            Description = "";
            HexString = "";
            IsDarkColor = isDarkColor;
            A = 255;
            R = 255;
            G = 255;
            B = 255;
        }

        internal ExcelColor(KnownColor color, bool isDarkColor)
        {
            Index = (int)color;
            IndexNPOI = (short)(Index + 7);
            Name = color.ToString();
            Description = "";
            HexString = "";
            IsDarkColor = isDarkColor;
            object[] attributeArray = color.GetType().GetField(Name).GetCustomAttributes(false);
            if (attributeArray != null && attributeArray.Length > 0 && attributeArray[0] is System.ComponentModel.DescriptionAttribute attribute)
            {
                if (!string.IsNullOrWhiteSpace(attribute.Description))
                {
                    string[] values = attribute.Description.Split(' ');
                    if (values.Length > 0)
                    {
                        Description = values[0];
                    }
                    if (values.Length > 1)
                    {
                        HexString = values[1];
                    }
                }
            }
            A = 255;
            R = int.Parse(HexString.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
            G = int.Parse(HexString.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
            B = int.Parse(HexString.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);

        }
        #endregion

        #region 公有方法

        #region 获取所有已知颜色
        public static List<ExcelColor> GetKnownColor()
        {
            List<ExcelColor> colorArray = new List<ExcelColor>();
            colorArray.Add(Black);
            colorArray.Add(White);
            colorArray.Add(Red);
            colorArray.Add(BrightGreen);
            colorArray.Add(Blue);
            colorArray.Add(Yellow);
            colorArray.Add(Pink);
            colorArray.Add(Turquoise);
            colorArray.Add(DarkRed);
            colorArray.Add(Green);
            colorArray.Add(DarkBlue);
            colorArray.Add(DarkYellow);
            colorArray.Add(Violet);
            colorArray.Add(Teal);
            colorArray.Add(Gray25Percent);
            colorArray.Add(Gray50Percent);
            colorArray.Add(Periwinkle);
            colorArray.Add(PlumPlus);
            colorArray.Add(Ivory);
            colorArray.Add(LiteTurquoise);
            colorArray.Add(DarkPurple);
            colorArray.Add(Coral);
            colorArray.Add(OceanBlue);
            colorArray.Add(IceBlue);
            colorArray.Add(DarkBluePlus);
            colorArray.Add(PinkPlus);
            colorArray.Add(YellowPlus);
            colorArray.Add(TurquoisePlus);
            colorArray.Add(VioletPlus);
            colorArray.Add(DarkRedPlus);
            colorArray.Add(TealPlus);
            colorArray.Add(BluePlus);
            colorArray.Add(SkyBlue);
            colorArray.Add(LightTurquoise);
            colorArray.Add(LightGreen);
            colorArray.Add(LightYellow);
            colorArray.Add(PaleBlue);
            colorArray.Add(Rose);
            colorArray.Add(Lavender);
            colorArray.Add(Tan);
            colorArray.Add(LightBlue);
            colorArray.Add(Aqua);
            colorArray.Add(Lime);
            colorArray.Add(Gold);
            colorArray.Add(LightOrange);
            colorArray.Add(Orange);
            colorArray.Add(BlueGray);
            colorArray.Add(Gray40Percent);
            colorArray.Add(DarkTeal);
            colorArray.Add(SeaGreen);
            colorArray.Add(DarkGreen);
            colorArray.Add(OliveGreen);
            colorArray.Add(Brown);
            colorArray.Add(Plum);
            colorArray.Add(Indigo);
            colorArray.Add(Gray80Percent);
            return colorArray;
        }
        #endregion

        #region 转换成RGB颜色
        public System.Drawing.Color ToRGBColor()
        {
            System.Drawing.Color color = System.Drawing.Color.FromArgb(A, R, G, B);
            return color;
        }
        #endregion

        #endregion


        #region Excel索引色
        /// <summary>
        /// Excel索引色。
        /// </summary>
        public enum KnownColor
        {
            /// <summary>
            /// 黑色
            /// </summary>
            [System.ComponentModel.Description("黑色 #000000")]
            Black = 1,
            /// <summary>
            /// 白色
            /// </summary>
            [System.ComponentModel.Description("白色 #FFFFFF")]
            White = 2,
            /// <summary>
            /// 红色
            /// </summary>
            [System.ComponentModel.Description("红色 #FF0000")]
            Red = 3,
            /// <summary>
            /// 鲜绿色
            /// </summary>
            [System.ComponentModel.Description("鲜绿色 #00FF00")]
            BrightGreen = 4,
            /// <summary>
            /// 蓝色
            /// </summary>
            [System.ComponentModel.Description("蓝色 #0000FF")]
            Blue = 5,
            /// <summary>
            /// 黄色
            /// </summary>
            [System.ComponentModel.Description("黄色 #FFFF00")]
            Yellow = 6,
            /// <summary>
            /// 粉红色
            /// </summary>
            [System.ComponentModel.Description("粉红色 #FF00FF")]
            Pink = 7,
            /// <summary>
            /// 青绿色
            /// </summary>
            [System.ComponentModel.Description("青绿色 #00FFFF")]
            Turquoise = 8,
            /// <summary>
            /// 深红色
            /// </summary>
            [System.ComponentModel.Description("深红色 #800000")]
            DarkRed = 9,
            /// <summary>
            /// 绿色
            /// </summary>
            [System.ComponentModel.Description("绿色 #008000")]
            Green = 10,
            /// <summary>
            /// 深蓝色
            /// </summary>
            [System.ComponentModel.Description("深蓝色 #000080")]
            DarkBlue = 11,
            /// <summary>
            /// 深黄色
            /// </summary>
            [System.ComponentModel.Description("深黄色 #808000")]
            DarkYellow = 12,
            /// <summary>
            /// 紫罗兰
            /// </summary>
            [System.ComponentModel.Description("紫罗兰 #800080")]
            Violet = 13,
            /// <summary>
            /// 青色
            /// </summary>
            [System.ComponentModel.Description("青色 #008080")]
            Teal = 14,
            /// <summary>
            /// 灰-25%
            /// </summary>
            [System.ComponentModel.Description("灰色25% #C0C0C0")]
            Gray25Percent = 15,
            /// <summary>
            /// 灰-50%
            /// </summary>
            [System.ComponentModel.Description("灰色50% #808080")]
            Gray50Percent = 16,
            /// <summary>
            /// 海螺色
            /// </summary>
            [System.ComponentModel.Description("海螺色 #9999FF")]
            Periwinkle = 17,
            /// <summary>
            /// 梅红色+
            /// </summary>
            [System.ComponentModel.Description("梅红色+ #993366")]
            PlumPlus = 18,
            /// <summary>
            /// 象牙色
            /// </summary>
            [System.ComponentModel.Description("象牙色 #FFFFCC")]
            Ivory = 19,
            /// <summary>
            /// 浅青色
            /// </summary>
            [System.ComponentModel.Description("浅青绿 #CCCCFF")]
            LiteTurquoise = 20,
            /// <summary>
            /// 深紫色
            /// </summary>
            [System.ComponentModel.Description("深紫色 #660066")]
            DarkPurple = 21,
            /// <summary>
            /// 珊瑚红
            /// </summary>
            [System.ComponentModel.Description("珊瑚红 #FF8080")]
            Coral = 22,
            /// <summary>
            /// 海蓝色
            /// </summary>
            [System.ComponentModel.Description("海蓝色 #0066CC")]
            OceanBlue = 23,
            /// <summary>
            /// 冰蓝色
            /// </summary>
            [System.ComponentModel.Description("冰蓝色 #CCCCFF")]
            IceBlue = 24,
            /// <summary>
            /// 深蓝色+
            /// </summary>
            [System.ComponentModel.Description("深蓝色+ #000080")]
            DarkBluePlus = 25,
            /// <summary>
            /// 粉红色+
            /// </summary>
            [System.ComponentModel.Description("粉红色+ #FF00FF")]
            PinkPlus = 26,
            /// <summary>
            /// 黄色+
            /// </summary>
            [System.ComponentModel.Description("黄色+ #FFFF00")]
            YellowPlus = 27,
            /// <summary>
            /// 青绿色+
            /// </summary>
            [System.ComponentModel.Description("青绿色+ #00FFFF")]
            TurquoisePlus = 28,
            /// <summary>
            /// 紫罗兰+
            /// </summary>
            [System.ComponentModel.Description("紫罗兰+ #800080")]
            VioletPlus = 29,
            /// <summary>
            /// 深红色+
            /// </summary>
            [System.ComponentModel.Description("深红色+ #800000")]
            DarkRedPlus = 30,
            /// <summary>
            /// 青色+
            /// </summary>
            [System.ComponentModel.Description("青色+ #008080")]
            TealPlus = 31,
            /// <summary>
            /// 蓝色+
            /// </summary>
            [System.ComponentModel.Description("蓝色+ #0000FF")]
            BluePlus = 32,
            /// <summary>
            /// 天蓝色
            /// </summary>
            [System.ComponentModel.Description("天蓝色 #00CCFF")]
            SkyBlue = 33,
            /// <summary>
            /// 浅青绿
            /// </summary>
            [System.ComponentModel.Description("浅青绿 #CCFFFF")]
            LightTurquoise = 34,
            /// <summary>
            /// 浅绿色
            /// </summary>
            [System.ComponentModel.Description("浅绿色 #CCFFCC")]
            LightGreen = 35,
            /// <summary>
            /// 浅黄色
            /// </summary>
            [System.ComponentModel.Description("浅黄色 #FFFF99")]
            LightYellow = 36,
            /// <summary>
            /// 淡蓝色
            /// </summary>
            [System.ComponentModel.Description("淡蓝色 #99CCFF")]
            PaleBlue = 37,
            /// <summary>
            /// 玫瑰红
            /// </summary>
            [System.ComponentModel.Description("玫瑰红 #FF99CC")]
            Rose = 38,
            /// <summary>
            /// 浅紫色
            /// </summary>
            [System.ComponentModel.Description("浅紫色 #CC99FF")]
            Lavender = 39,
            /// <summary>
            /// 茶色
            /// </summary>
            [System.ComponentModel.Description("茶色 #FFCC99")]
            Tan = 40,
            /// <summary>
            /// 浅蓝色
            /// </summary>
            [System.ComponentModel.Description("浅蓝色 #3366FF")]
            LightBlue = 41,
            /// <summary>
            /// 水绿色
            /// </summary>
            [System.ComponentModel.Description("水绿色 #33CCCC")]
            Aqua = 42,
            /// <summary>
            /// 酸橙色
            /// </summary>
            [System.ComponentModel.Description("酸橙色 #99CC00")]
            Lime = 43,
            /// <summary>
            /// 金色
            /// </summary>
            [System.ComponentModel.Description("金色 #FFCC00")]
            Gold = 44,
            /// <summary>
            /// 浅橙色
            /// </summary>
            [System.ComponentModel.Description("浅橙色 #FF9900")]
            LightOrange = 45,
            /// <summary>
            /// 橙色
            /// </summary>
            [System.ComponentModel.Description("橙色 #FF6600")]
            Orange = 46,
            /// <summary>
            /// 蓝灰色
            /// </summary>
            [System.ComponentModel.Description("蓝灰色 #666699")]
            BlueGray = 47,
            /// <summary>
            /// Gray-40%
            /// </summary>
            [System.ComponentModel.Description("灰色40% #969696")]
            Gray40Percent = 48,
            /// <summary>
            /// 深青色
            /// </summary>
            [System.ComponentModel.Description("深青色 #003366")]
            DarkTeal = 49,
            /// <summary>
            /// 海绿色
            /// </summary>
            [System.ComponentModel.Description("海绿色 #339966")]
            SeaGreen = 50,
            /// <summary>
            /// 深绿色
            /// </summary>
            [System.ComponentModel.Description("深绿色 #003300")]
            DarkGreen = 51,
            /// <summary>
            /// 橄榄绿
            /// </summary>
            [System.ComponentModel.Description("橄榄绿 #333300")]
            OliveGreen = 52,
            /// <summary>
            /// 褐色
            /// </summary>
            [System.ComponentModel.Description("褐色 #993300")]
            Brown = 53,
            /// <summary>
            /// 梅红色
            /// </summary>
            [System.ComponentModel.Description("梅红色 #993366")]
            Plum = 54,
            /// <summary>
            /// 靛蓝色
            /// </summary>
            [System.ComponentModel.Description("靛蓝色 #333399")]
            Indigo = 55,
            /// <summary>
            /// 灰-80%
            /// </summary>
            [System.ComponentModel.Description("灰色80% #333333")]
            Gray80Percent = 56,

        }
        #endregion
    }
}
