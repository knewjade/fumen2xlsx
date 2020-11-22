import common.tetfu.Tetfu
import common.tetfu.TetfuPage
import common.tetfu.common.ColorConverter
import common.tetfu.common.ColorType
import core.field.FieldFactory
import core.mino.MinoFactory
import org.apache.commons.cli.DefaultParser
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream
import java.nio.charset.StandardCharsets
import java.nio.file.Files
import java.nio.file.Paths
import java.util.*

fun extractV115(text: String): String? {
    val regex = "v115@[a-zA-Z0-9?+/]+"
    return regex.toRegex().find(text)?.value?.let { Tetfu.removePrefixData(it) }
}

fun main() {
    val url =
//        "v115@HhglBeBtAeQ4CeglywBtR4RphlwwzhQ4RpJeAgHAhQ?4FeglBeR4EeglywQ4i0RpAeglwwxSwhgHg0RpJeAAA/gAtH?eBtFeglg0AtwwgWwwR4ilAth0gWAeQ4zhJeAAA/gAtHeBtD?eglBeg0AtwhxwilRpg0hlwwzhQpwhJeAAAChQ4Deg0DeR4C?eh0glywQ4glRpwhxSwhwwhlAeRpJeAAAChQ4DeglDeR4Ceg?lBtxwgWQ4i0APglBtwwhHxhg0JeAAAJhglBeBtAeQ4AeRpg?lxwgWBtR4whQphlwwhHxhQ4JeAAAChQ4FeglBeR4CeRpgly?wQ4i0RpAeglwwxSwhgHg0JeAAAIhAtAeR4Beg0BeBtR4wwg?Wwwg0RpAtwhQaxhgWQag0RpJeAAABhAtHeBtDeg0ilAtywR?4g0glwSwhhHwwQ4whh0JeAAABhAtHeBtBeg0BeilAtwhxwg?0RpglxhxSwwh0QpwSJeAAABhAtHeBtDeglRpg0AtwhxwilR?pg0hlwwwhQLxhJeAAAEhQ4Beg0FeR4Aei0AtAeywQ4glyhw?SAeAtwwhlKeAAAEhQ4Deg0DeR4AeRpi0xwgWQ4glwSQpzhw?wRLglJeAAADhAtHeBtBeg0RpilAtwwgWwwg0RpAPzhgWQag?0JeAAAKhAtAeR4Beg0RpBtR4whxwg0RpAtxhxSwwg0KeAAA?AhRpBeAtBeglBeRpAeBtBeglywR4Ati0hlwwQpQ4hHxhKeA?AAAhRpg0DeglBeRpg0BeQ4AeglxwglQpg0AtAeR4Aeglwwh?WwhwSAeAtQ4JeAAAAhQ4FeglBeR4RpAeAtg0glxwglAtwSQ?pBtg0AeglwwhWQLwhAtQag0JeAAAChQ4DeglBeBtR4h0Aeg?lxwgWBtQ4g0RpAeglwwhHhWg0QawSJeAAAzgQ4IeR4CeAtD?eglQ4BeBtDeglywAtQah0RpAeglwwxShHAewSQpJeAAAChQ?4DeQLBeBtR4AeRpQLgWxwAtAeQ4g0RpAeQLwwwhxSwhi0Je?AAAAhQ4FeglBeR4RpBtAeglywQ4APQpg0BtAeglwwxShHg0?RpJeAAA/gAtEehlAeBtDeRpglg0AtwwgWwwR4RpglAth0gW?AeAtgWxSwhJeAAA/gAtCehlCeBtDeglRpg0AtwhxwR4glxh?g0hlwwR4whhWQaJeAAA9gQ4EehlBeR4BeBtAeglBeg0Q4wh?xwAtglAeRpg0hlwwyhQaLeAAA9gQ4IeR4BeBtAeilQaQ4yw?AtglAeRpRag0wwxSwhQaRpJeAAA/gAtBeglQ4DeBtBeglR4?BeQaAtywhlQ4RpRag0wwxSwhQaxhJeAAA9ghlDeQ4DeglBe?BtR4Beg0glwhxwglwSQ4Rpg0hlwwwhQahHRpJeAAA/gAtHe?BtBeRpQ4hlAtglywQpAPR4glg0Atg0wwxShHQ4glJeAAAAh?AtFeRpBtBei0AeRpAeywR4g0glwhxSwhwwBtQLLeAAAAhAt?Deh0RpBtDeg0AeRpAtwhxwR4g0glhWxSwwQ4whBPglJeAAA?ChRpBeAtg0DeRpAeBti0whxwR4gWAPhWxSwwR4glAPglJeA?AA/gRpEeAtg0AeRpDeBti0xwglAtQ4AtgWQLyhwwBtglhWJ?eAAA+gAtHeBtR4Beg0AehlAtQ4Atglxwg0RpgWQLwhhWwwh?0QpQagWJeAAA+gAtHeBtR4Beg0RpAeAtQpQ4ywg0RpglxSh?Hwwg0AeglLeAAAAhAtFeRpBtBeg0AeR4QpQaAtywg0R4glx?ShHwwg0AeRaglJeAAAAhAtFeRpBtBeg0glQ4AeRpAtwhxwg?0glR4hWxSwwg0AeglQLKeAAA+gAtDeh0BeBtR4Beg0CeAtR?4whxwg0glRphWxSwwglAPAeAPQpJeAAA+gAtg0BeQ4DeBtg?0BeR4CeAtglg0ywwhAPRpxShHwwglAPglRpJeAAA9gh0EeA?tBeg0AeR4BeBtBeg0QpwhxwgWAtglRpxSQLwhwwRLglRpJe?AAAChRpBewhglDeRpAeg0whAPgWAtywQ4Qpg0whAPglxSww?Q4Qpglg0whJeAAA/gRpEewhglAeRpDeg0whglBtxwglAtQ4?g0BeglBtwwBtg0LeAAA+gg0Hegli0BeRpQ4AeglAtwhglxw?RpQ4g0AeglxhwwyhQLg0JeAAA2gg0Ieg0FeglAeh0BeRpQ4?AeglgWAtywRpR4AeAPhWwwwhhWgHQ4JeAAAChRpg0BeglBe?gHAeRpg0Q4AeglBtQaxwh0R4AeglhWwwwhhWgHQ4JeAAA/g?RpCeg0BeglAeRpCeg0Q4AeglBtxwglQpAeR4AeglBtwwhWQ?LgHQ4JeAAA1gQ4IeR4CeAtDeglQ4BeBtBeRpQaglxwAtAeh?0whQpRawwxhQLgHg0JeAAADhQ4CeBtwDglBeR4Rpg0AtwSQ?LywSph0AtQLglwwxSxhJeAAAFhQ4BeRpglBeBtR4g0Rpglx?wgWBtQ4hlg0AeglwwhHhWJeAAAFhglBeAtRpBeilg0BtRpy?wQ4Aeg0AthWQawhwwRpRaJeAAA5gglIeglDeAtRpBehlAeg?0BtRpywQ4wwg0AtwShWwhywAtKeAAA/gglCeRpCeAtglCew?hQpAeg0BthlxwQpR4g0AtwShWwhywg0KeAAA/gglRpFeAtg?lRpDeg0BtAeglxwglAtQ4g0AtwSQaxhwwBtg0KeAAAFhQ4A?eh0AeglBeBtR4g0QpwhQaglxwBtQ4g0QawhRawwyhQLJeAA?A9gwhEeRpBewhglDeRpAeg0whglBtywQ4Aeg0QLAPhlAtww?RpRaJeAAA9gwhBeRpEewhglAeRpDeg0AeglBtxwglAtQ4g0?BeglBtwwBtg0KeAAA6gAtDeQ4CeBtDeR4BeAtg0BewhhlAt?glxwg0RpwhQawhhWwwh0QpwSJeAAA9gQ4EeRpBeR4DeQpQa?Aeglg0Q4wSAtxwwhgWhlh0QawSAtwwwSQLxhJeAAA9gQ4Be?RpEeR4AeRpDeglQaQ4BtxwglQahlRag0BtwwhWwhQaJeAAA?9ghlDeQ4DeglDeR4RpQaglAtwhglxwQ4RpRag0xhwwyhQaJ?eAAADhQ4DeilBeR4RpAtAeBtxwgWQ4Rpg0Atg0BtwwhHhWJ?eAAADhQ4DeAtg0glBeR4RpBtg0glxwgWQ4RpCtAeglwwhHh?WJeAAA/gQ4EehlBeR4BeBtAeglRpg0Q4xwgWBtglQpQaAeA?tg0wwhHhWJeAAABhAtBeglQ4DeBtBeglR4RpQaAtywglAew?wRpRag0wwxSwhQaJeAAA/ghlDeQ4DeglBeBtR4Rpg0glwhx?wglwSQ4Rpg0hlwwwhQahHJeAAA/gh0CeQ4BeRpg0DeR4AeR?pAtglAtywAeglyhQaBtwhgWRLJeAAA9gh0EeQ4Beg0RpDeR?4Aeg0QpwSgWAtxwgWQ4glhWQLwhBtwwRLglJeAAAEhQ4Beh?0AtRpBeR4BegWAtRpxwgWQ4glAewhgWyhwwRLglJeAAAAhA?tDeh0BeBtR4Beg0AeRpAtR4xwgWg0glyShWwhwwRLglJeAA?AAhAtg0BeQ4DeBtg0BeR4AeRpAtglg0ywwhAPRpxShHwwgl?APglJeAAAGhAti0AeR4BeBtRpg0QpwhxwgWAtglRpxSQLwh?wwRLglJeAAA/gh0EeAtBeg0AeR4BeBtRpg0R4ywAtgWxSQL?whhHwwglhWJeAAA8gAtDeQ4CeBtDeR4BeAtg0RpBPglQ4yw?glRpAPwhhHwhwwhlJeAAA/gQ4BeRpEeR4AeRpBejlQ4wSAt?xwwhhlyhgHwSAtwwglKeAAADhAtCeglQ4AeRpBtBeg0glQ4?AeRpwhglxwg0hlAegHwhhWwwg0KeAAABhglRpDeAtBeglRp?Beg0BtR4hlwhxwg0glAeQ4hWxSwwg0KeAAABhAtEeRpAeBt?R4Beg0RpglAtwwAtglxwg0glgWglwhQahWwwg0KeAAADhAt?DeR4RpBtBeg0R4glQpQaAtywg0RaglxShHwwg0KeAAABhAt?FehlBtR4Beg0RpglAtR4whxwg0QpAPQLhWxSwwg0KeAAA/g?BtAeQ4DeywBtR4glRpg0wwxhwSwhQ4glRpg0BtBPxDhlgHJ?eAAA/gBtAeQ4DeywBtR4ilg0wwxSwhgHQ4glRpGegWwhQpJ?eAAA0gAtHeBtFeQ4AeAtywzhR4Rpwhglg0QailQ4xSgHBeg?0glAexSJeAAA0gAtHeBtGewhAtwhxwR4i0whRawwR4ilQaw?hEeAPgWBewhJeAAA0gAtHeBtFeglAeAtwhxwR4ilg0RawwR?4zhwwEegWBeg0wwJeAAA0gAtHeBtFeQ4AeAtwhxwzhR4Raw?wilh0AtQ4CeglEegHJeAAA0gAtHeBtDeglCeAtwhxwili0R?awwzhQ4Aeg0CegWBeR4gHKeAAAsgglIeglIehlEeQ4whxwz?hTpwwi0glAtRpwSAtCeg0wDwSAtKeAAA4gglHehlDeQ4xwQ?pzhRpR4wwBtg0BtRpTeAAApgglIeglIehlHeQ4ywwhQaxhR?pR4wwEtRpTeAAA9ggWBeBtDewhzwBtilwhywi0glRpwhGeR?awhJeAAAAhBtDewhQ4ywBtglRpwhR4wwhlg0glRpGeglQaw?SKeAAAAhBtAewhCeQ4ywBtwhglRpR4wwhlg0whglQpglFeQ?aAeglQaJeAAAAhBtAewhCeQ4ywBtwhilR4wwhlg0AeglRpG?egWwhQpJeAAA2gg0Iei0DeQ4ywBtzhR4wwRaAtQpilCeRpg?HQaglxSJeAAAtgg0Ieg0Heh0EeQ4xwQaBtzhR4wwxSgWAti?lTeAAAqgg0Ieg0Heh0BegHEeQ4ywgWAtzhR4wwxSBtilTeA?AAsgglIeglGegHAehlEeQ4QaxwBt1hwwRpBti0GegWAeg0J?eAAA4gglHehlDeQ4xwQpBtzhR4wwxhBti0TeAAApgglIegl?IehlHeQ4ywAtglzhR4wwxhglAti0TeAAA9ggWBeBtDewhzw?Bti0whywRpilg0whEeglCeQLJeAAAAhBtCeglAeQ4ywBtil?g0R4wwRazhwwEegWBeg0wwJeAAA2gRpHeRpAeglCeQ4ywil?i0R4wwBeyhQag0CewSwhAtBehHJeAAAzgRpHeRpDeglCeQ4?xwglQahli0R4wwyhQaxhg0TeAAAAhi0DeAtglxwAtg0zhBt?wwBtQpwhilCeAtAPQaQpglLeAAAAhi0DeQ4ywAtg0zhR4ww?hWAPhlRpEeQawSgWRpJeAAA2gg0Iei0DeQ4ywzhRpR4wwCP?BtRpCegWAegWAeAtwhwSJeAAAtgg0Ieg0Heh0EeQ4xwQazh?RpR4wwCPBtRpTeAAAqgg0Ieg0Heh0BegHEeQ4ywQLyhRpR4?wwBPglBtRpTeAAA+ggHAeBtDewhQ4QaxwBti0yhwwilRpg0?whFeQpQaAPwhJeAAAAhBtDewhQ4ywBtRpg0whR4wwhWglRp?g0GewSQpg0KeAAAAhBtAewhCeQ4ywBtwhRpg0R4wwhWglwh?RpwwFewhgHAewwJeAAAAhBtAewhCeQ4ywBtwhi0R4wwhWgl?AeRpg0GeQpwSgHJeAAABhBtAeg0BeRpywBti0RpglwhwSwh?AewhR4glQLglgWBeQaBtKeAAA+gR4Feg0R4ywzhg0ilwwRa?AtgWh0AehWAeRpAexSKeAAA+gR4BeRpg0BeR4ywRpi0glRL?wwBtyhQLDewSwhAtBPKeAAA+gR4Beg0BeRpR4ywi0RpglRL?wwAtwhgWQLxhTeAAA+gR4FewhR4ywAti0whglRLwwAtgWSp?gWDeAtAPQaQpAewhJeAAA+gR4FewhR4ywAtRpg0whglRLww?BtRpg0GewSQpg0KeAAA+gR4CewhCeR4ywAtwhRpg0glRLww?BtwhRpwwFewhgHAewwJeAAA+gR4CewhCeR4ywAtwhi0glRL?wwBtAeRpg0GeQpwSgHJeAAA9gRpBeBtAeg0BeRpywBti0gl?RLwwxhAewhR4DeAPAeQaBtKeAAA3gQ4FeAtBeR4DeBtywQ4?0hAeQpwwhlg0APhlgWRpCeg0QLwDKeAAA+gAtBezhBeBtxw?gWR4i0AtQaQpwwBeilg0FeAPgWLeAAA3gQ4FeAtBeR4DeBt?ywQ4zhAtQaQpwwyhwwh0DeglAegWBeg0JeAAA9gwhCeBtDe?whQ4xwgWBti0whQpQ4wwBPilg0whwSAtAewhQpglBegHJeA?AA9gwhCeBtDewhQ4ywBtglRpAeR4wwhlg0glRpDexSAthlK?eAAA9gwhCeBtDewhQ4ywBtilAeR4wwhlg0glRpGegWwhQpJ?eAAA9gwhCeBtDewhQ4ywBtRpg0AeR4wwhWglRpg0DeglAeg?HAeRpJeAAA9gwhCeBtDewhQ4ywBti0AeR4wwhWglRpg0GeQ?pwSgHJeAAA1gAtGewhBtGewhAtywR4i0AeRpwwRpilg0AeQ?pAPAegWAeglxSKeAAA+gR4FewhR4whxwBthlwhwwhlwwRpB?tglwhQawSQpAeRpgWAeglwhJeAAA+gR4BeRpBeglR4ywRpi?lg0RawwBtyhEewSwhAtAegWQaJeAAA+gR4DeglRpR4ywilR?pg0RawwAtwhgWxhQaTeAAA+gR4FewhR4ywAtilwhg0RawwB?tglwhAegWDeAtBPRpwhJeAAA+gR4FewhR4ywAtglRpwhg0R?awwBtglRpGeglQawSKeAAA+gR4CewhCeR4ywAtwhglRpg0R?awwBtwhglQpglFeQaAeglQaJeAAA+gR4CewhCeR4ywAtwhi?lg0RawwBtAeglRpGegWwhQpJeAAA9gRpBeBtCeglRpywBti?lg0RawwxhAewhR4DeAPAeQaQ4AtwSJeAAABhBtCeglh0ywB?tilQpAeQpwwxSxhQ4wwg0QpwSQeAAAChBtAeQ4AeilywBtR?4glRpg0whwSyhwwgHBei0AexDKeAAA2gAtHeBtEewhRpAty?wi0whRpR4wwhWglQawhAeBthHAPCewhJeAAA2gAtHeBtDeg?lAeRpAtwhxwilg0TpwwzhwwEegWBeg0wwJeAAAChBtAeQ4A?eglRpwhxwBtR4glQpQaglwwyhQaQ4glQLwDi0BehHJeAAA4?gQ4IeR4BewhilywQ4Rpwhgli0wwBeQpAPwhhWBehHBtAewh?JeAAA9gzhDeQ4AeilxwgWBtR4gli0wwBPBtg0EeQpQaAPAe?QaJeAAA4gRpHeRpBewhilywR4AtxhzwR4AtwSwhEexSAtAe?whJeAAA1gRpHeRpEewhilxwglAtQ4Atwhgli0wwDtUeAAA9?gwhAeR4FewhQ4AtglxwAtRpg0whglRawwBtRpwwwhglAegH?AeAtAeAPg0wwJeAAA9gwhAeR4FewhR4ywAti0AeglRLwwBt?Rpg0GeQpwSgHJeAAA9gwhAeR4FewhR4ywAtglRpAeg0Raww?BtglRpAegWAeg0CeRaKeAAA9gwhAeR4FewhR4ywAtilAeg0?RawwBtglRpGegWwhQpJeAAA9gwhAeR4FewhR4ywBthlAeg0?RawwRpBtglEeQaQpAewSQaJeAAA6gglGeilBeQ4Btywzhg0?Q4xSwwi0RpQaQ4AegHAexSg0QpwhJeAAAugglIeglIehlCe?Q4BtywwhQaxhR4BtwwCtRpTeAAArgglIeglIehlAegWDeQ4?BtxwQpzhR4BtwwBtg0RpTeAAA2gglGeglAeglFeQ4Atglyw?zhR4hlwwi0RpTeAAAChhlBewhQ4BtQpxwglRpwhQ4wwhlww?g0glRpwhEeh0AewSgWJeAAA4gg0Iei0BeQ4BtywzhR4Btww?xhhlwhEexSAPAeQaJeAAAvgg0Ieg0Heh0CeQ4BtxwQazhR4?BtwwxSAPhlTeAAA0gg0Iei0BegHCeQ4BtywQLyhR4BtwwxS?ilTeAAAsgg0Ieg0Heh0FeQ4gWAtywzhQ4whhWwwRpilTeAA?A6gglEegHAeilBeQ4BtQaxwzhR4hWwwRpi0GegWAeg0JeAA?AugglIeglIehlCeQ4BtywwhQaxhR4BtwwxhAth0TeAAArgg?lIeglIehlAegWDeQ4BtxwQpzhR4Btwwxhi0TeAAA2gglGeg?lAeglFeQ4AtglywzhR4hlwwRpi0TeAAA4gRpHeRpBewhQ4B?tQpxwi0whQ4wwhlwwilg0whEeQawSBeQLJeAAA1gRpHeRpE?ewhQ4BtxwglQph0whR4BtwwRaglg0UeAAA4gg0Iei0BeQ4A?twhglxwzhR4xhwwilQpglHeQpglJeAAAvgg0Ieg0Heh0CeQ?4BtxwQazhR4BtwwCPRpTeAAA0gg0Iei0BegHCeQ4BtywQLy?hR4BtwwBPglRpTeAAAsgg0Ieg0Heh0FeQ4gWAtywzhQ4whh?WwwilRpTeAAA4gRpFegHAeRpAeglAeQ4BtQaxwilg0R4hWw?wzhg0EegWBeRpJeAAA1gRpHeRpDeglAeQ4BtxwglQahlg0R?4BtwwhWwhQag0TeAAA/gh0EewhSpglxwR4AtwhSpQawwR4A?tglwhAeQLhlCeAtgHQLJeAAA5gg0Iei0AezhywR4AtRpBPg?lwwR4AtRpwhAegWCeAPAtQaJeAAAwgg0Ieg0Heh0BezhxwQ?aR4AtRpilwwxhgWAtTeAAA1gg0Iei0BegHBezhywwhQ4AtR?pilwwxhBtTeAAAtgg0Ieg0Heh0EexhQLwhywR4AtRpCPwwR?4BtTeAAA+gAtAeQ4whEeBtR4QaxwRpg0AtxhRLwwglRpg0x?SgWAeilAeglg0JeAAA+gAtAeR4EeBtR4ywi0AtwSwhhHwwg?lRpg0GeQpwSgHJeAAADhBtAewhilQ4ywBtwhglQaQpBewwi?0whAeRpAewwhWwSQpwhJeAAADhBtAewhglRpQ4ywBtwhglR?pR4wwhlg0AeglQawSQeAAA5gQ4Deg0DeR4Begli0ywQ4Rpg?lzhwwBeQpglDewDBeAtgWQaJeAAA5gRpCeg0DeRpBegli0x?wgWR4AtglQLyhwwBeBtGeAPLeAAA2gRpFeg0AeRpEegli0x?wglAtQ4AtglQLyhwwDtTeAAA5gg0Iei0AeyhgWglxwR4Atg?lAPglBewwR4BtAegWAeRpOeAAAwgg0Ieg0Heh0BezhxwQaR?4AtilRpwwxhgWAtTeAAA1gg0Iei0BegHBezhywwhQ4AtilR?pwwxhBtTeAAAtgg0Ieg0Heh0EexhQLwhywR4AthlAPxSwwR?4BtTeAAA+gAtBegHAezhBtR4Qaxwi0AtR4xSwwilg0gWEeg?lAeAPKeAAADhBtAewhi0Q4ywBtwhQpQag0R4wwyhAeRpAew?SAtDewhJeAAA9gwhEehlBewhQ4BtywglRpwhR4BtwwglgWQ?pglgWwSQ4AewDAeAth0QaJeAAA5gRpBewhEeRpBewhQ4Bty?wi0AeR4BtwwBeglg0FeAPhHKeAAA2gRpEewhBeRpEewhQ4B?txwglQph0AeR4BtwwRaglg0TeAAA7gglGeilAeyhgWglxwR?4AtglQpg0RpwwR4BtglQpwDAeg0AegWAeAtKeAAAvgglIeg?lIehlBezhywQ4wwAtRpi0ywglAtTeAAA3gglGeilAegWCez?hxwQpR4AtRpi0ywBtTeAAAsgglIeglIehlEezhQpxwR4AtR?pCtwwR4BtTeAAA3gAtEewhBeglAtEewhRpglywi0whRpywi?lg0gWwSR4gHAeglAeAPKeAAA7gglGeilAe0hxwR4Atwwh0R?awwR4BtQaAeQaAPQpAegWAeAtKeAAAvgglIeglIehlBezhy?wQ4wwAti0RpywglAtTeAAA3gglGeilAegWCezhxwQpR4Ati?0RpywBtTeAAAsgglIeglIehlEezhQpxwR4Ath0AtxhwwR4B?tTeAAA9gwhBeAtg0EewhRpAtywR4AtwhRpAtAewwR4BtwhA?eAPRaOeAAADhBtAewhh0R4ywBtwhwwR4xSwwilwhwwAegWx?hAeglAeAPwhJeAAADhBtAewhRpg0Q4ywBtwhRpg0R4wwhWg?lAegHh0wSAtOeAAA5gQ4CewhEeR4BewhilywQ4Rpwhgli0w?wBeQpglwhhHAeQaAegWBtQaJeAAA5gRpBewhEeRpBewhilx?wgWR4AtAegli0wwBeBtGeAPLeAAA2gRpEewhBeRpEewhilx?wglAtQ4AtAegli0wwDtTeAAA9gRpBeglBeBtAeRpglRaywB?tAexhgWAtQ4wwi0QaBeQ4whCeAPg0JeAAA/gglRpBeBtAei?lRpywBthWxhwwQ4wwhlg0TeAAA9gRpg0DeBtAeRpi0ywBtx?hQagWAtQ4wwhWglGeglAegHJeAAA9gg0BeRpBeBtAei0Rpy?wBthWQLwhR4wwhWglTeAAABhR4CewhilR4ywAtwhAPTpg0w?wBewhAeRpxDg0AeglAewhJeAAAAhwhCeBtAeilwhQ4ywBtg?lRpwhBewwh0wwCewhAewhAeAPAewwJeAAABhR4CewhglRpR?4ywAtwhglRpwwh0wwBewhglQawSQaAeQaAeAtAeQLJeAAAA?hwhCeBtAeglRpwhQ4ywBtglRpwhBewwh0wwCewhAewhAeAP?AewwJeAAABhR4CewhilR4ywAtwhglh0wwRpwwBewhhWAeww?QpAPAeAtAeQLJeAAA/gzhBeQ4AeilBtywR4gli0xSwwRpg0?DexSAeQaQpQaJeAAA6gQ4IeR4whilBtywQ4whglg0xwRpww?QpAPwhIewhJeAAA4gAtHeBtBeQ4AezhAtxwgWR4ili0wwBP?g0glBegHAeg0CeQaJeAAA9gglFeBtAegl0hxwBthlR4Raww?g0Qag0gWR4AeQpwSAexSg0JeAAA9gwhFeBtAewhh0R4ywBt?Qag0R4RpwwhWglwhQawDDeglAegHJeAAABhR4Cewhi0R4yw?AtwhglQpg0ilwwBewhglwSAeglxSAeglAewhJeAAAAhwhCe?BtAei0whQ4ywBtRpg0whBewwhlwhCeQaAeQ4AegWAeQaJeA?AA9gwhFeBtAewhi0Q4ywBtwhRpwwR4wwhWglgWAeQpQaPeA?AA4gAtHeBtBeQ4AezhAtywR4wwh0ilwwRaQ4QawSQpglAew?DAewhQpKeAAA9gg0CeR4Dei0R4whxwhlBtxhxSwwQpAPglA?eAthWEeglJeAAA/gglAeR4DeilR4ywRpgWAtxhhHwwg0RpG?eRpAtJeAAA/gglAeR4BeRpilR4ywRpBtQawhhHwwi0GehHK?eAAA9gg0CeR4BeRpi0R4ywRpBtQawhhHwwglRaGeglAegHJ?eAAA9gwhFeBtAewhRpg0Q4ywBtQLRpg0BewwglRawhAPglg?0AeQ4NeAAABhR4CewhRpg0R4ywAtwhglQpg0ilwwBewhQag?0AeAPAewDAeglAewhJeAAA9gwhFeBtAewhglRpQ4ywBtwhg?lRpBewwh0wwwhBPgWAeQ4AeAPAewwJeAAA9gwhFeBtAewhi?lQ4ywBtAeglRpR4wwhlg0AegWwhQpPeAAAAhwhCeBtAeRpg?0whQ4ywBtglQpg0whR4wwhWglQag0QpgWCeglAegHJeAAA7?gQ4BeAtFeR4BtzhywQ4Athlwhi0whQaQpAeAPgHQaAewDg0?gWRpJeAAAAhwhAeR4CeilwhR4xwgWAtglQaQpwhi0wwxSAe?whQpwhDewhwSJeAAAygg0Ieg0Heh0RpzhywAtRpQ4g0glRL?wwBtAeBtQaglAegHMeAAAvgg0Ieg0Heh0BegHRpzhywgWRp?R4ilwwhWTeAAA3gg0Ieh0DeRpzhQaxwAtRpR4glBPwwBtTe?AAAxgglIeglIehlRpxhQLwhywAtRpR4CewwBtDegWAeg0Me?AAAugglIeglIehlAegWAeRpzhxwQpAtRpR4i0wwhlTeAAA5?gglGeglAeglCeRpyhQaywAtRpR4g0BtwwBtTeAAA7gQ4Deg?0DeR4ili0QpxwQ4glBtwhSawwRpAewDwSAtBegHAeQaQpJe?AAA7gRpCeg0DeRpili0xwgWAtglxhQLwhR4wwxSBeBPR4Be?whwSJeAAA4gRpFeg0AeRpCeili0xwglwhglxhQLwhR4wwxh?TeAAAAhwhAeR4CeglRpwhQ4AtglxwAtglRpQLg0RpwwBthl?AewhxDg0MeAAA+gg0CeR4Cegli0R4ywAtglxhAewhBPwwBt?CeQaAeQpwSMeAAA7gQ4Beg0FeR4gli0BtywQ4glQLyhxSww?RpEexSAeQaQpJeAAAygg0Ieg0Heh0zhR4xwgWAtglAPglR4?RpwwxSAegWCeRpAewhwSJeAAAvgg0Ieg0Heh0BegHzhR4yw?gWilR4RpwwhWTeAAA3gg0Ieh0DezhR4QaxwAtilR4xSwwBt?TeAAAAhwhAeR4Cei0xhQ4ywAtRpg0whCPwwBtwhQpAewhgl?xSMeAAAxgglIeglIehlzhR4ywAtj0Q4BPwwBtxSg0QagWRp?MeAAAugglIeglIehlAegWAezhR4xwQpAti0R4RpwwhlTeAA?A5gglGeglAeglCezhQ4zwAti0R4xhwwBtTeAAABhilCezhg?lQ4QpxwAtRph0AtywBtRpgHAeg0wSAtMeAAA9gwhDeR4Cew?hi0R4ywAtwhRpg0CewwBtgWAeQpAeAPAewDMeAAABhilCez?hglQ4ywAtwwh0RpBewwBtQawSRpwhAeQ4MeAAA9gwhDeR4C?ewhRpg0R4ywAtwhRpg0CewwBtwhBeQpQaAewDMeAAA9gwhD?eR4CewhglRpR4ywAtAeglRpg0RawwBtAeglAPgHgWAeg0Me?AAA9gwhDeR4CewhilR4ywAtAeglRpg0RawwBtAegWwhQpPe?AAA7gQ4AewhGeR4whilBtywQ4Aegli0xSwwRpBexSg0AegH?AeQaQpJeAAA9gwhDeR4CewhilR4xwgWAtAegli0RpwwxSEe?RpAewhwSJeAAAAhwhAeR4CeRpg0whR4ywAtglQpg0whglRL?wwBtQah0whAPxSMeAAA7gQ4BeAtFeR4BtzhywQ4Ath0wwgl?RLwwRpAehHwwDeQaQpJeAAABhAtAeR4Bei0BtR4wwgWwwQp?Qag0AtzhgWQLRpAegHgWBeglRaJeAAABhAtAeR4BeRpg0Bt?R4ywRpg0AtwSwhhHwwglwSQpg0QeAAA3ghlIeglBeQ4Aezh?wwglBtR4i0wwwhQpAPwSAtQ4AegHBewwRpiWJeAAA3gh0Ee?AtBeg0EeBtR4AtQpzhAtR4SpwwglQLglBegHQpgWglwSglL?eAAA"
        "v115@VgwhGeRpwhGeRpwhh0Eehlwhg0DeR4A8glA8g0AeBt?R4B8glB8BeBtF8AeJ8AeF8JeNFYeAlfnBCRrDfEVGs6AlvF?LBFIEfETYdzBlvs2A4lAAAzgQ4GexwQ4GexwQ4hlEeh0Q4g?lDexhAeg0TeAAPAAkgzhglEeBtilCei0BtFeQ4g0RpFeR4R?pGeQ4XeAAP2AlfnBCxsDfEVGs6AlvFLBFIEfETYk2Alvs2A?EqDfET4cBClvs2AGFEfET4dBBlfnBCxZAAA7gRLFexwRLFe?xwfeAAP2AlfnBCxsDfEVGs6AlvFLBFIEfETYk2Alvs2AEqD?fET4cBClvs2AGFEfET4dBBlfnBCRbAAAvhBtqQhAUNKSAVa?W3A4XPDCGKjRA1gJ+BxXnQB2HMSAVaW3AZAAAAAAPAA4gT4?g0EeBti0CeilBtxwDewhglBexwCeRpwhNeAAPdAQIKvDll2?TASIvBElCyTAVaW3AYblRAVaW3AZAAAAzgRpi0EeRpR4g0F?eR4HeywHewwQeAAPqAQIKvDll2TASoW9AzXHDBQY85AQ1gR?ASo78A1no2AlsCSASY9tCv/AAAzgzhglE8BtilF8BtH8ywH?8wwG8JeAAPqAQIKvDll2TASoW9A0XHDBQY85AQ1gRASo78A?1no2AlsCSASYlFDM+AAAzgilRpE8glh0RpF8g0wwH8g0xwH?8wwG8JeAAPpAQIKvDll2TASoW9A1XHDBQYkRBOx78AwngHB?FbcRAzB88AQuHgCsAAAAzgi0RpE8Btg0RpF8BtH8ywH8wwG?8JeAAPpAQIKvDll2TASoW9A2XHDBQYkRBus78AwngHBFbcR?AzB88AQurgCqAAAAzgQ4zhE8R4ilF8Q4glH8ywH8wwG8JeA?APpAQIKvDll2TASoW9A3XHDBQYMOBu178AwngHBFbcRAzB8?8AQOstCpAAAAzgQ4hlRpE8R4glRpF8Q4glH8ywH8wwG8JeA?APpAQIKvDll2TASoW9A4XHDBQYMOBOx78AwngHBFbcRAzB8?8AQ+ytCsAAAAzgi0RpE8hlg0RpF8glwwH8glxwH8wwG8JeA?APpAQIKvDll2TASoW9A5XHDBQY0KBuy78AwngHBFbcRAzB8?8AQOMgCqAAAAzgBti0E8zhg0F8BtH8ywH8wwG8JeAAPpAQI?KvDll2TASouABwXHDBQY0KBOo78AwngHBFbcRAzB88AQeNP?C6AAAAzgBtglRpE8ilRpF8BtH8ywH8wwG8JeAAPpAQIKvDl?l2TASouABxXHDBQY0KBOo78AwngHBFbcRAzB88AQ+KWC6AA?AAzgRpR4glE8RpilF8R4H8ywH8wwG8JeAAPpAQIKvDll2TA?SouAByXHDBQYcHBOr78AwngHBFbcRAzB88AQOMgCzAAAAzg?Q4hlRpE8R4glRpF8BtH8Q4BtH8glG8JeAAPpAQIKvDll2TA?SouABzXHDBQYEEBO078AwngHBFbcRAzB88AQHUWCvAAAAzg?hli0E8zhg0F8glwwH8glxwH8wwG8JeAAPpAQIKvDll2TASo?uAB0XHDBQYsABO078AwngHBFbcRAzB88AQeNPCsAAAAzgQ4?zhE8R4ywF8hlH8Q4glwwH8glG8JeAAPpAQIKvDll2TASouA?B1XHDBQYsABO078AwngHBFbcRAzB88AQC+tCpAAAAzgzhAt?E8ywBtF8hlH8wwglAtH8glG8JeAAPpAQIKvDll2TASouAB2?XHDBQYsABO078AwngHBFbcRAzB88AQijxCpAAAAzgi0RpE8?ywRpF8BtH8wwBtH8g0G8JeAAPpAQIKvDll2TASouAB3XHDB?QYsABO078AwngHBFbcRAzB88AQXegCqAAAAzgQ4ywAtE8R4?wwBtF8hlH8Q4glAtH8glG8JeAAPpAQIKvDll2TASouAB4XH?DBQYU9Au178AwngHBFbcRAzB88AQiLuC0AAAAzgg0ywAtE8?i0BtF8RpH8RpAtH8wwG8JeAAPpAQIKvDll2TASouAB5XHDB?QYU9Au178AwngHBFbcRAzB88AwmzPC0AAAAzgzhAtE8i0Bt?F8RpH8RpAtH8g0G8JeAAPpAQIKvDll2TASoGEBwXHDBQYU9?Au178AwngHBFbcRAzB88AwmzPCpAAAAzgQ4ywglE8R4ilF8?BtH8Q4BtH8wwG8JeAAPpAQIKvDll2TASoGEBxXHDBQYU9Au?178AwngHBFbcRAzB88AQHUWC0AAAAzgQ4zhE8R4ilF8BtH8?Q4BtH8glG8JeAAPpAQIKvDll2TASoGEByXHDBQYU9Au178A?wngHBFbcRAzB88AQHUWCpAAAAzgQ4i0glE8R4ilF8Q4wwH8?xwg0H8wwG8JeAAPpAQIKvDll2TASoGEBzXHDBQYU9Au178A?wngHBFbcRAzB88AQOstCqAAAAzgQ4ilAtE8R4whBtF8Q4wh?H8glwhAtH8whG8JeAAPpAQIKvDll2TASoGEB0XHDBQYU9Au?178AwngHBFbcRAzB88AwtjFDsAAAAzgQ4ywAtE8R4g0BtF8?Q4g0H8h0AtH8wwG8JeAAPpAQIKvDll2TASoGEB1XHDBQY85?Au178AwngHBFbcRAzB88AQvjFD0AAAAzgzhAtE8i0BtF8Q4?g0H8R4AtH8Q4G8JeAAPpAQIKvDll2TASoGEB2XHDBQY85Au?178AwngHBFbcRAzB88AwsPFDpAAAA"
//        "v115@vhKRQJKJJUGJvMJTNJ+DJdAndFndKnFKnFKJXhywHe?wwLeAgH"
    main("-t $url -l 8 -n 10 -c default -o output/test.xlsx".split(" "))
//    main("-t $url -l 8 -n 10 -c mipi".split(" "))
}

fun main(commands: List<String>) {
    val options = XlsxOptions.create()
    val settingParser = XlsxSettingParser(options, DefaultParser())
    val settingsOptional = settingParser.parse(commands)
    val settings = settingsOptional.get()

    val colorTheme = settings.colorTheme

    val properties = properties(colorTheme)

    val data = extractV115(settings.fumen)
    val minoFactory = MinoFactory()
    val colorConverter = ColorConverter()
    val pages = Tetfu(minoFactory, colorConverter).decode(data)

    val workbook = XSSFWorkbook()
    val maxHeight = settings.height?.takeIf { 0 < it } ?: pages.maxOf { it.field.usingHeight }
    val horizontalNum = settings.horizontalNum.takeIf { 0 < it } ?: error("Invalid horizontal number")
    val visibleComment = settings.comment

    val cellColors = CellColors(workbook, properties)
    val gridSheet = GridSheet.create(workbook, cellColors, settings.cellSize, horizontalNum, maxHeight, visibleComment)

    Main(minoFactory, colorConverter).run(gridSheet, pages, horizontalNum)

    FileOutputStream(settings.output).use {
        workbook.write(it)
    }
}

class Main(private val minoFactory: MinoFactory, private val colorConverter: ColorConverter) {
    fun run(gridSheet: GridSheet, pages: List<TetfuPage>, horizontalNum: Int) {
        pages.forEachIndexed { index, page ->
            val fx = index % horizontalNum
            val fy = index / horizontalNum

            if (fx == 0) {
                gridSheet.setPage(fy, "${index + 1}")
            }

            val operation = page.colorType.takeIf { ColorType.isMinoBlock(it) }?.let {
                val field = FieldFactory.createField(24)
                val piece = colorConverter.parseToBlock(it)
                val mino = minoFactory.create(piece, page.rotate)
                field.put(mino, page.x, page.y)
                Operation(it, field, page.isLock)
            }

            gridSheet.draw(fx, fy, page.field, operation, page.comment)
        }
    }
}

private fun properties(colorTheme: String): Properties {
    val properties = Properties()
    properties.load(
        Files.newBufferedReader(
            Paths.get(String.format("theme/${colorTheme}.properties")),
            StandardCharsets.UTF_8
        )
    )
    return properties
}