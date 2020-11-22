import common.tetfu.common.ColorType
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import java.awt.Color
import java.util.*

class CellColors(private val workbook: Workbook, private val properties: Properties) {
    private val boarderColor = toColor(properties.getProperty("Border"))

    private val normalMap: Map<ColorType, XSSFCellStyle> = mutableMapOf<ColorType, XSSFCellStyle>().apply {
        put(ColorType.T, solid(toColor(properties.getProperty("T.normal"))))
        put(ColorType.I, solid(toColor(properties.getProperty("I.normal"))))
        put(ColorType.O, solid(toColor(properties.getProperty("O.normal"))))
        put(ColorType.S, solid(toColor(properties.getProperty("S.normal"))))
        put(ColorType.Z, solid(toColor(properties.getProperty("Z.normal"))))
        put(ColorType.L, solid(toColor(properties.getProperty("L.normal"))))
        put(ColorType.J, solid(toColor(properties.getProperty("J.normal"))))
        put(ColorType.Gray, solid(toColor(properties.getProperty("Gray.normal"))))
        put(ColorType.Empty, solid(toColor(properties.getProperty("Empty"))))
    }.toMap()

    private val clearMap: Map<ColorType, XSSFCellStyle> = mutableMapOf<ColorType, XSSFCellStyle>().apply {
        put(ColorType.T, solid(toColor(properties.getProperty("T.clear"))))
        put(ColorType.I, solid(toColor(properties.getProperty("I.clear"))))
        put(ColorType.O, solid(toColor(properties.getProperty("O.clear"))))
        put(ColorType.S, solid(toColor(properties.getProperty("S.clear"))))
        put(ColorType.Z, solid(toColor(properties.getProperty("Z.clear"))))
        put(ColorType.L, solid(toColor(properties.getProperty("L.clear"))))
        put(ColorType.J, solid(toColor(properties.getProperty("J.clear"))))
        put(ColorType.Gray, solid(toColor(properties.getProperty("Gray.clear"))))
    }.toMap()

    private val pieceMap: Map<ColorType, XSSFCellStyle> = mutableMapOf<ColorType, XSSFCellStyle>().apply {
        put(ColorType.T, solid(toColor(properties.getProperty("T.piece"))))
        put(ColorType.I, solid(toColor(properties.getProperty("I.piece"))))
        put(ColorType.O, solid(toColor(properties.getProperty("O.piece"))))
        put(ColorType.S, solid(toColor(properties.getProperty("S.piece"))))
        put(ColorType.Z, solid(toColor(properties.getProperty("Z.piece"))))
        put(ColorType.L, solid(toColor(properties.getProperty("L.piece"))))
        put(ColorType.J, solid(toColor(properties.getProperty("J.piece"))))
    }.toMap()

    fun normal(type: ColorType): XSSFCellStyle {
        return normalMap.get(type) ?: error("invalid normal type")
    }

    fun clear(type: ColorType): XSSFCellStyle {
        return clearMap.get(type) ?: error("invalid clear type")
    }

    fun normalOrClear(type: ColorType, cleared: Boolean): XSSFCellStyle {
        return if (cleared) {
            clear(type)
        } else {
            normal(type)
        }
    }

    fun piece(type: ColorType): XSSFCellStyle {
        return pieceMap.get(type) ?: error("invalid piece type")
    }

    private fun solid(c: Color): XSSFCellStyle {
        val color = XSSFColor(c)
        val boarderColor = XSSFColor(boarderColor)
        return (workbook.createCellStyle() as XSSFCellStyle).apply {
            setFillForegroundColor(color)
            setFillBackgroundColor(color)
            setFillPattern(FillPatternType.SOLID_FOREGROUND)

            borderTop = CellStyle.BORDER_THIN
            borderBottom = CellStyle.BORDER_THIN
            borderLeft = CellStyle.BORDER_THIN
            borderRight = CellStyle.BORDER_THIN

            setTopBorderColor(boarderColor)
            setBottomBorderColor(boarderColor)
            setLeftBorderColor(boarderColor)
            setRightBorderColor(boarderColor)
        }
    }

    companion object {
        fun toColor(code: String): Color {
            return Color.decode(code)
        }
    }
}

