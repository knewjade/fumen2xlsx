import common.tetfu.common.ColorType
import common.tetfu.field.ColoredField
import core.field.Field
import core.field.FieldFactory
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.*

class Operation(val colorType: ColorType, val minoField: Field, val locked: Boolean) {
    fun exists(x: Int, y: Int): Boolean {
        return minoField.exists(x, y)
    }
}

class GridSheet(
    private val sheet: XSSFSheet, private val cellColors: CellColors,
    private val rowHeight: Float, private val fieldHeight: Int, private val visibleComment: Boolean
) {
    fun setPage(fy: Int, comment: String) {
        val top = fy * fieldHeight + fy * (if (visibleComment) 2 else 1) + 1
        row(top).createCell(0).also {
            it.setCellValue(comment)
        }
    }

    fun draw(fx: Int, fy: Int, field: ColoredField, operation: Operation?, comment: String?) {
        val left = fx * 10 + fx + 2
        val top = fy * fieldHeight + fy * (if (visibleComment) 2 else 1) + 1

        val f = FieldFactory.createField(24)
        (0 until 24).forEach { y ->
            (0 until 10).forEach { x ->
                if (field.getColorType(x, y) != ColorType.Empty) {
                    f.setBlock(x, y)
                }
            }
        }

        operation?.let {
            if (it.locked) {
                f.merge(it.minoField)
            }
        }

        val maxY = fieldHeight - 1
        (maxY downTo 0).forEach { y ->
            val row = row(top + (maxY - y))
            val cleared = f.getBlockCountOnY(y) == 10

            (0 until 10).forEach { x ->
                val style: XSSFCellStyle = operation?.takeIf { it.exists(x, y) }?.let {
                    cellColors.piece(operation.colorType)
                } ?: run {
                    val type = field.getColorType(x, y)
                    cellColors.normalOrClear(type, cleared)
                }

                row.createCell(left + x).also {
                    it.cellStyle = style
                }
            }
        }

        comment?.let { str ->
            if (visibleComment) {
                val y = top + maxY + 1
                sheet.addMergedRegion(CellRangeAddress(y, y, left, left + 9))
                row(y, flag = false).createCell(left).also {
                    it.setCellValue(str)
                    it.cellStyle.setAlignment(HorizontalAlignment.LEFT)
                }
            }
        }
    }

    private fun row(y: Int, flag: Boolean = true): XSSFRow {
        return sheet.getRow(y) ?: sheet.createRow(y).also {
            if (flag) {
                it.heightInPoints = rowHeight
            }
        }
    }

    fun ceil(x: Int, y: Int): XSSFCell? {
        return row(y).getCell(x)
    }

    companion object {
        fun create(
            workbook: XSSFWorkbook, cellColors: CellColors,
            cellSize: Double, horizontalNum: Int, fieldHeight: Int, visibleComment: Boolean
        ): GridSheet {
            val width = 256 * cellSize
            val height = cellSize * (49 / 9.3)
            val sheet = workbook.createSheet().also {
                val max = horizontalNum * 11 + 2
                (0 until max).forEach { x ->
                    it.setColumnWidth(x, width.toInt())
                }
            }

            return GridSheet(sheet, cellColors, height.toFloat(), fieldHeight, visibleComment)
        }
    }
}