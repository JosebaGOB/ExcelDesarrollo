using ClosedXML.Excel;

namespace ConsoleApplication18.Excel
{



    public class Estilo
    {
        public Fuente Fuente { get; set; }
        public Bordes Bordes { get; set; }
        public Relleno Relleno { get; set; }

        public void AplicarEstilo(IXLStyle style)
        {
            if (Fuente != null)
            {
                Fuente.AplicarEstilo(style.Font);
            }

            if (Bordes != null)
            {
                Bordes.AplicarEstilo(style.Border);
            }

            if (Relleno != null)
            {
                Relleno.AplicarEstilo(style.Fill);
            }
        }
    }


    public class Relleno
    {
        public XLColor ColorFondo { get; set; }
        public XLColor ColorPatron { get; set; }
        public XLFillPatternValues TipoPatron { get; set; }

        public Relleno()
        {
            ColorFondo = XLColor.NoColor;
            ColorPatron = XLColor.NoColor;
            TipoPatron = XLFillPatternValues.Solid;
        }

        public void AplicarEstilo(IXLFill style)
        {
            style.BackgroundColor = ColorFondo;
            style.PatternColor = ColorPatron;
            style.PatternType = TipoPatron;
        }

    }


    public class Fuente
    {
        public bool Negrita { get; set; }
        public bool Italica { get; set; }
        public XLFontUnderlineValues Subrayado { get; set; }
        public bool Tachado { get; set; }
        public XLFontVerticalTextAlignmentValues AlineacionVertical { get; set; }
        public bool Sombra { get; set; }
        public double TamanyoFuente { get; set; }
        public XLColor Color { get; set; }
        public string NombreFuente { get; set; }
        public XLFontCharSet FontCharSet { get; set; }


        public Fuente(string nombreFuente = "Calibri", double tamanyoFuente = 11, XLColor fontColor = null)
            : this(nombreFuente, tamanyoFuente, fontColor, false, false, XLFontCharSet.Default)
        {
        }

        public Fuente(string nombreFuente, double tamanyoFuente, XLColor fontColor, bool negrita, bool italica,
            XLFontCharSet fontCharSet)
        {
            Negrita = negrita;
            Italica = italica;
            TamanyoFuente = tamanyoFuente;
            Color = (fontColor == null) ? XLColor.Black : fontColor;
            NombreFuente = nombreFuente;
            FontCharSet = fontCharSet;
            Subrayado = XLFontUnderlineValues.None;
            Tachado = false;
            AlineacionVertical = XLFontVerticalTextAlignmentValues.Baseline;
            Sombra = false;
        }

        public void AplicarEstilo(IXLFont style)
        {
            style.Bold = Negrita;
            style.Italic = Italica;
            style.Underline = Subrayado;
            style.Strikethrough = Tachado;
            style.VerticalAlignment = AlineacionVertical;
            style.Shadow = Sombra;
            style.FontSize = TamanyoFuente;
            style.FontColor = Color;
            style.FontName = NombreFuente;
            style.FontCharSet = FontCharSet;
        }
    }


    public enum PosicionBordes
    {
        Todos,
        Externo,
        Ninguno,
        Superior,
        Inferior,
        Izquierdo,
        Derecho
    }


    public class Bordes
    {
        public PosicionBordes Posicion { get; set; }
        public XLBorderStyleValues Estilo { get; set; }
        public XLColor Color { get; set; }

        public Bordes()
        {
            Posicion = PosicionBordes.Ninguno;
            Estilo = XLBorderStyleValues.Medium;
            Color = XLColor.Black;
        }

        public void AplicarEstilo(IXLBorder border)
        {
            if (Estilo == XLBorderStyleValues.None || Posicion == PosicionBordes.Ninguno)
            {
                return;
            }

            switch (Posicion)
            {
                case PosicionBordes.Todos:
                    border.InsideBorder = Estilo;
                    border.InsideBorderColor = Color;
                    border.RightBorder = Estilo;
                    border.RightBorderColor = Color;
                    border.LeftBorder = Estilo;
                    border.LeftBorderColor = Color;
                    border.BottomBorder = Estilo;
                    border.BottomBorderColor = Color;
                    border.TopBorder = Estilo;
                    border.TopBorderColor = Color;
                    break;
                case PosicionBordes.Externo:
                    border.RightBorder = Estilo;
                    border.RightBorderColor = Color;
                    border.LeftBorder = Estilo;
                    border.LeftBorderColor = Color;
                    border.BottomBorder = Estilo;
                    border.BottomBorderColor = Color;
                    border.TopBorder = Estilo;
                    border.TopBorderColor = Color;
                    break;
                case PosicionBordes.Superior:
                    border.TopBorder = Estilo;
                    border.TopBorderColor = Color;
                    break;
                case PosicionBordes.Inferior:
                    border.BottomBorder = Estilo;
                    border.BottomBorderColor = Color;
                    break;
                case PosicionBordes.Izquierdo:
                    border.LeftBorder = Estilo;
                    border.LeftBorderColor = Color;
                    break;
                case PosicionBordes.Derecho:
                    border.RightBorder = Estilo;
                    border.RightBorderColor = Color;
                    break;
            }
        }
    }
}