import common.tetfu.Tetfu
import common.tetfu.common.ColorConverter
import core.mino.MinoFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.assertj.core.api.Assertions.assertThat
import org.junit.jupiter.api.Test
import java.util.*


class XlsxMainTest {
    private val minoFactory = MinoFactory()
    private val colorConverter = ColorConverter()
    private val main = MainRunner(minoFactory, colorConverter)

    @Test
    fun case1W4() {
        val properties = Properties()
        properties.load(javaClass.classLoader.getResource("test.properties")?.openStream())

        val workbook = XSSFWorkbook()
        val cellColors = CellColors(workbook, properties)

        val horizontalNum = 4
        val sheet = GridSheet.create(workbook, cellColors, 2.0, horizontalNum, 5, false)

        val data = "vhGSQJJHJWSJUIJXGJTJJVBJ"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)

        main.run(sheet, pages, horizontalNum)

        assertThat(sheet.ceil(0, 1)?.stringCellValue).isEqualTo("1")
        assertThat(sheet.ceil(0, 7)?.stringCellValue).isEqualTo("5")

        assertThat(sheet.ceil(44, 1)).isNotNull()
        assertThat(sheet.ceil(46, 1)).isNull()

        assertThat(sheet.ceil(5, 5)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xff.toByte(), 0xaa.toByte(), 0x33.toByte())
        assertThat(sheet.ceil(5, 11)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xb3.toByte(), 0x77.toByte(), 0x24.toByte())
    }

    @Test
    fun case1W5() {
        val properties = Properties()
        properties.load(javaClass.classLoader.getResource("test.properties")?.openStream())

        val workbook = XSSFWorkbook()
        val cellColors = CellColors(workbook, properties)

        val horizontalNum = 5
        val sheet = GridSheet.create(workbook, cellColors, 2.0, horizontalNum, 5, false)

        val data = "vhGSQJJHJWSJUIJXGJTJJVBJ"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)

        main.run(sheet, pages, horizontalNum)

        assertThat(sheet.ceil(0, 1)?.stringCellValue).isEqualTo("1")
        assertThat(sheet.ceil(0, 7)?.stringCellValue).isEqualTo("6")

        assertThat(sheet.ceil(44, 1)).isNotNull()
        assertThat(sheet.ceil(46, 1)).isNotNull()

        assertThat(sheet.ceil(5, 5)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xff.toByte(), 0xaa.toByte(), 0x33.toByte())
        assertThat(sheet.ceil(5, 11)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xb3.toByte(), 0x77.toByte(), 0x24.toByte())
    }

    @Test
    fun case2CommentOff() {
        val properties = Properties()
        properties.load(javaClass.classLoader.getResource("test.properties")?.openStream())

        val workbook = XSSFWorkbook()
        val cellColors = CellColors(workbook, properties)

        val horizontalNum = 4
        val sheet = GridSheet.create(workbook, cellColors, 2.0, horizontalNum, 5, false)

        // コメント付き
        val data = "vhGSQYFAooMDEPBAAAJHJWSJUIJXGJTJJVBJ"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)

        main.run(sheet, pages, horizontalNum)

        assertThat(sheet.ceil(0, 1)?.stringCellValue).isEqualTo("1")
        assertThat(sheet.ceil(0, 7)?.stringCellValue).isEqualTo("5")

        assertThat(sheet.ceil(5, 5)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xff.toByte(), 0xaa.toByte(), 0x33.toByte())
        assertThat(sheet.ceil(5, 11)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xb3.toByte(), 0x77.toByte(), 0x24.toByte())

        assertThat(sheet.ceil(46, 1)).isNull()
    }

    @Test
    fun case2CommentOn() {
        val properties = Properties()
        properties.load(javaClass.classLoader.getResource("test.properties")?.openStream())

        val workbook = XSSFWorkbook()
        val cellColors = CellColors(workbook, properties)

        val horizontalNum = 4
        val sheet = GridSheet.create(workbook, cellColors, 2.0, horizontalNum, 5, true)

        // コメント付き
        val data = "vhGSQYFAooMDEPBAAAJHJWSJUIJXGJTJJVBJ"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)

        main.run(sheet, pages, horizontalNum)

        assertThat(sheet.ceil(0, 1)?.stringCellValue).isEqualTo("1")
        assertThat(sheet.ceil(0, 8)?.stringCellValue).isEqualTo("5")

        assertThat(sheet.ceil(5, 5)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xff.toByte(), 0xaa.toByte(), 0x33.toByte())
        assertThat(sheet.ceil(5, 12)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0xb3.toByte(), 0x77.toByte(), 0x24.toByte())
    }

    @Test
    fun case3() {
        val properties = Properties()
        properties.load(javaClass.classLoader.getResource("test.properties")?.openStream())

        val workbook = XSSFWorkbook()
        val cellColors = CellColors(workbook, properties)

        val horizontalNum = 4
        val sheet = GridSheet.create(workbook, cellColors, 2.0, horizontalNum, 10, false)

        val data = "zgwhIewhH8AewhH8AewhH8AeI8KepIJvhAAgH"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)

        main.run(sheet, pages, horizontalNum)

        assertThat(sheet.ceil(2, 0)).isNull()
        assertThat(sheet.ceil(2, 1)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0x15.toByte(), 0x15.toByte(), 0x00.toByte())

        assertThat(sheet.ceil(2, 6)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0x00.toByte(), 0xb3.toByte(), 0xb3.toByte())
        assertThat(sheet.ceil(2, 7)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0x00.toByte(), 0xd9.toByte(), 0xd9.toByte())
        assertThat(sheet.ceil(11, 7)?.cellStyle?.fillBackgroundXSSFColor?.rgb)
                .containsExactly(0x00.toByte(), 0xff.toByte(), 0xff.toByte())

        assertThat(sheet.ceil(2, 11)).isNull()
    }

    @Test
    fun detect1() {
        val data = "zgwhIewhH8AewhH8AewhH8AeI8KepIJvhAAgH"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)
        assertThat(main.detectHeight(pages)).isEqualTo(5)
    }

    @Test
    fun detect2() {
        val data = "vhGSQJJHJWSJUIJXGJTJJVBJ"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)
        assertThat(main.detectHeight(pages)).isEqualTo(5)
    }

    @Test
    fun detect3() {
        val data = "vhGSQYZAlvs2ARoeRA1gJ+BxXnQB2HMSA1d85ARAAA?AJHJWSJUIJXGJVBJTJJ3gQpHeSpQ4GexhQ4BtAexwAexhg0?Q4glBtxwAei0Q4ilJeAgWxAlvs2ARoeRA1gJ+BxXnQB2HMS?A1dcHBzXHDBwPjRA1dMOBFYHDBQRsRA1d85ARAAAAvhNO/X?ZAlvs2AR4gRA1gJ+BxXnQB2HMSA1d85ARAAAA62Ip+Iz9Iv?3IM4INmmNrmNwm11mdAndFnFFJAgH1gh0whBeAtEeg0xhBt?AeQ4hlAeg0AewhAtxwQ4glFexwQ4TeAgWxAlvs2AR4gRA1g?J+BxXnQB2HMSA1dcHBzXHDBwPjRA1dMOBFYHDBQRsRA1d85?ARAAAA"
        val pages = Tetfu(minoFactory, colorConverter).decode(data)
        assertThat(main.detectHeight(pages)).isEqualTo(10)
    }
}