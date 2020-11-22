import entry.CommandLineWrapper
import entry.common.SettingParser
import exceptions.FinderParseException
import org.apache.commons.cli.CommandLineParser
import org.apache.commons.cli.Options
import java.util.*

class XlsxSettingParser(options: Options?, parser: CommandLineParser?) :
    SettingParser<XlsxSettings>(options, parser) {
    @Throws(FinderParseException::class)
    override fun parse(wrapper: CommandLineWrapper): Optional<XlsxSettings> {
        val fumen = wrapper.getStringOption(XlsxOptions.Fumen.optName()).orElse(null) ?: error("Cannot get fumen")
        val cellSize = wrapper.getDoubleOption(XlsxOptions.CellSize.optName()).orElse(2.0)
        val height = wrapper.getIntegerOption(XlsxOptions.Line.optName()).orElse(null)
        val horizontalNum = wrapper.getIntegerOption(XlsxOptions.HorizontalNum.optName()).orElse(5)
        val colorTheme = wrapper.getStringOption(XlsxOptions.ColorTheme.optName()).orElse("default")
        val comment = wrapper.getBoolOption(XlsxOptions.Comment.optName()).orElse(true)
        val output = wrapper.getStringOption(XlsxOptions.Output.optName()).orElse("output/output.xlsx")
        val settings = XlsxSettings(fumen, cellSize, height, horizontalNum, colorTheme, comment, output)
        return Optional.of(settings)
    }
}