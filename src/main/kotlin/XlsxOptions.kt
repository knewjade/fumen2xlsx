import entry.common.option.OptionBuilder
import entry.common.option.SingleArgOption
import org.apache.commons.cli.Options

enum class XlsxOptions(private val optionBuilder: OptionBuilder) {
    Fumen(SingleArgOption.full("t", "tetfu", "v115@~", "Specify tetfu data")),
    CellSize(SingleArgOption.full("s", "size", "double", "Specify cell size")),
    Line(SingleArgOption.full("l", "line", "integer", "'Set field height")),
    HorizontalNum(SingleArgOption.full("n", "num", "integer", "'Change horizontal field count")),
    ColorTheme(SingleArgOption.full("c", "color", "string", "'Change color theme")),
    Comment(SingleArgOption.full("C", "comment", "yes or no", "'Show comment flag")),
    Output(SingleArgOption.full("o", "output", "path", "'Change output path")),
    ;

    fun optName(): String {
        return optionBuilder.longName
    }

    companion object {
        fun create(): Options {
            val allOptions = Options()
            for (options in values()) {
                val option = options.optionBuilder.toOption()
                allOptions.addOption(option)
            }
            return allOptions
        }
    }
}