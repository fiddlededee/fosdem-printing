@file:Suppress("Since15")

import common.normalizeImageDimensions
import converter.fodt.AsciidoctorOdAdapter
import converter.fodt.FodtConverter
import model.Image
import model.Length
import model.LengthUnit
import model.NoWriterNode
import model.Node
import model.OpenBlock
import model.Paragraph
import model.Span
import model.TRG
import model.TableCell
import org.asciidoctor.Asciidoctor
import org.asciidoctor.Asciidoctor.Factory
import org.asciidoctor.Options
import org.asciidoctor.SafeMode
import org.jsoup.Jsoup
import org.jsoup.nodes.Document
import org.languagetool.JLanguageTool
import org.languagetool.language.Russian
import org.languagetool.rules.spelling.SpellingCheckRule
import org.openqa.selenium.chrome.ChromeDriver
import org.openqa.selenium.chrome.ChromeOptions
import org.openqa.selenium.print.PageSize
import org.openqa.selenium.print.PrintOptions
import reader.HtmlNode
import reader.UnknownTagProcessing
import writer.OdtCustomWriter
import writer.OdtStyle
import writer.OdtStyleList
import writer.paragraphProperties
import writer.tableCellProperties
import writer.textProperties
import java.io.File
import java.lang.Thread.sleep
import java.nio.file.Files
import java.nio.file.Paths
import java.time.Duration
import java.util.*
import kotlin.io.path.Path
import kotlin.io.path.createDirectory
import com.helger.css.reader.CSSReader
import model.Col
import model.Heading
import model.Table
import model.Toc
import writer.tableProperties

version = "1.0.1"
val presentationFile = "fosdem-printing"

val spell = properties["spell"]

buildscript {
    repositories {
        mavenCentral()
    }
    dependencies {
        classpath("org.asciidoctor:asciidoctorj:3.0.0")
        classpath("org.jsoup:jsoup:1.17.1")
        classpath("com.google.guava:guava:21.0")
        classpath("org.languagetool:language-ru:5.6")
        classpath("org.seleniumhq.selenium:selenium-java:4.40.0")
        classpath("io.github.bonigarcia:webdrivermanager:5.8.0")
        classpath("org.apache.poi:poi:5.2.5")
        classpath("org.apache.poi:poi-ooxml:5.2.5")
        classpath("ru.fiddlededee:unidoc-publisher:0.9.2")
        classpath("com.helger:ph-css:8.1.1")
    }
}

val makePdf = tasks.register("makePDF") {
    group = "talk"
    description = "Creates PDF from HTML version"
    mustRunAfter(makeHtml)
    doLast {
        arrayOf("print-pdf&pdfSeparateFragments=false" to "no-animation-", "print-pdf" to "animation-").forEach {
            saveToPdf(it.first, it.second)
        }
    }
}

val makeHtml = tasks.register("makeHTML") {
    group = "talk"
    description = "Creates HTML version of presentation"
    doLast {
        if (File("output").exists()) File("output").deleteRecursively()
        Path("output").createDirectory()

        File("${project.projectDir}/$presentationFile.adoc")
            .run { AsciidocHtmlFactory().getHtmlFromFile(this) }
            .run {
                val title = Jsoup.parse(this).selectFirst("h1")?.wholeText()
                    ?: "Title undefined"
                println(title)
                this.jsoupParse().apply {
                    select("section:not(.title):not(.no-footer)")
                        .forEach { it.append(footerHtml(title)) }
                }.apply {
                    if (spell == null) return@apply
                    this.selectXpath("//p").forEach { it.text().langToolsCheck() }
                }
                    .toString()
            }.toFile("${project.projectDir}/output/$presentationFile-v$version.html")

        File("${project.projectDir}/reveal.js").copyRecursively(
            File("${project.projectDir}/output/reveal.js"),
            overwrite = true
        )
        File("${project.projectDir}/roboto-2014").copyRecursively(
            File("${project.projectDir}/output/roboto-2014"),
            overwrite = true
        )
        File("${project.projectDir}/images").copyRecursively(
            File("${project.projectDir}/output/images"),
            overwrite = true
        )
        File("${project.projectDir}/white-course.css").copyTo(
            File("${project.projectDir}/output/white-course.css"),
            overwrite = true
        )
        File("${project.projectDir}/version.txt").copyTo(
            File("${project.projectDir}/output/version.txt"),
            overwrite = true
        )
    }
}

val makeOd = tasks.register("makeOD") {
    group = "talk"
    description = "Creates OD (fodt) version of presentation with notes"
    doLast {
        File("output").mkdir()
        // tag::boiler-plate[]
        FodtConverter {
            html = AsciidocHtmlFactory()
                .getHtmlFromFile(File("${project.projectDir}/$presentationFile.adoc"), true)
            template = File("${project.projectDir}/template-1.fodt").readText()
            adaptWith(AsciidoctorOdAdapter)
            unknownTagProcessingRule = unknownTagProcessingRuleRevealJs()
            parse()
            // end::boiler-plate[]
            // tag::orchestration-rebuild-title[]
            ast().descendant { section ->
                section.sourceTagName == "section" &&
                        section.descendant { it is Heading && it.level == 1 }
                            .isNotEmpty()
            }.first().also { it.insertBefore(makeTitle(it)) }.remove() // <1>
            // end::orchestration-rebuild-title[]
            // tag::orchestration-other[]
            ast().descendant { it.roles.contains("notes") } // <1>
                .forEach { it.insertBefore(HorizontalLine()) }
            ast().descendant { it is Heading && it.level > 1 }
                .forEach {
                    it.insertBefore(
                        Paragraph().apply { roles("slide-finish") }
                    )
                } // <2>
            odtStyleList.add(odtStyles())
            odtStyleList.add(rougeStyles()) // <3>
            // end::orchestration-other[]
            // bad decision to take Reveal.js as source for HTML, but safety margin of the approach allows
            arrayOf("fa-lightbulb-o" to "Tip", "fa-warning" to "Warning")
                .forEach { mapUnit ->
                    ast().descendant { it.roles.contains(mapUnit.first) }.forEach {
                        it.insertBefore(Paragraph().apply { +mapUnit.second })
                        it.remove()
                    }
                }
            // Ha! Order matters, interesting to check OD specification
            ast().descendant { it is Table && (it.parent()?.roles?.contains("colist") ?: false) }
                .map { it as Table }.forEach {
                    it.children().first()
                        .apply { arrayOf(8F, 162F).forEach { insertBefore(Col(Length(it))) } }
                }
            File("${project.projectDir}/output/${presentationFile}-prenote.html").writeText(html())
            // tag::boiler-plate-2[]
            ast2fodt()
            File("${project.projectDir}/output/ast.yaml").writeText(ast().toYamlString())
            File("${project.projectDir}/output/$presentationFile-notes-v$version.fodt").writeText(fodt())
        }
        // end::boiler-plate-2[]
    }
}

val fodt2All = tasks.register<Exec>("Fodt2All") {
    group = "talk"
    description = "Converts notes from FODT (LibreOffice) to all formats, needs FODT file, created with ${makeOd.name}, needs LibreOffice and Kotlin installed"
    commandLine(
        "kotlin",
        "lo-kts-converter.main.kts",
        "-i",
        "output/fosdem-printing-notes-v1.0.1.fodt",
        "-f",
        "pdf,docx,odt"
    )

}


class AsciidocHtmlFactory {
    private val factory: Asciidoctor = Factory.create()
    fun getHtmlFromFile(file: File, sectnums: Boolean = false): String {
        return factory.convertFile(
            file,
            Options.builder().backend("html5").sourcemap(true)
                .apply {
                    if (sectnums)
                        attributes(org.asciidoctor.Attributes.builder().attribute("sectnums").build())
                }.safe(SafeMode.UNSAFE).toFile(false).standalone(true)
                .templateDirs(File("${project.projectDir}/templates"))
                .build()
        )
    }
}

fun String.jsoupParse(): Document {
    return Jsoup.parse(this)
}

fun String.println(): String {
    println(this)
    return this
}

fun String.toFile(name: String): String {
    File(name).writeText(this)
    return this
}


fun footerHtml(title: String): String {
    return """
            <div class="twd1">
                <div>$title</div>
                <img 
                    src="images/fosdem-logo.svg"  
                    alt="TechWriter Days #1" 
                    class="tw-days"/>
            </div>""".trimIndent()
}

object LangTools {
    private val langTool = JLanguageTool(Russian())
    var ruleTokenExceptions: Map<String, Set<String>> = mapOf()
    var ruleExceptions: Set<String> = setOf("")

    fun setSpellTokens(
        ignore: Array<String>,
        accept: Array<String> = arrayOf()
    ): LangTools {
        langTool.allActiveRules.forEach { rule ->
            if (rule is SpellingCheckRule) {
                rule.addIgnoreTokens(ignore.toList())
                rule.acceptPhrases(accept.toList())
            }
        }
        return this
    }

    fun check(text: String) {
        val errs = langTool.check(text).filterNot {
            (this.ruleTokenExceptions[it.rule.id]?.contains(text.substring(it.fromPos, it.toPos)) ?: false) or
                    ((this.ruleExceptions).contains(it.rule.id))
        }

        if (errs.isNotEmpty()) {
            var errorMessage = "Spell failed for:\n$text\n"
            errs.forEachIndexed { index, it ->
                errorMessage += "[${index + 1}] ${it.message}, ${it.rule.id} (${it.fromPos}:${it.toPos} " +
                        "- ${text.substring(it.fromPos, it.toPos)})\n"
            }
            println(errorMessage.split("\n").map { it.chunked(120) }.flatten().joinToString("\n"))
        }
    }
}

fun String.langToolsCheck() {
    LangTools
        .apply {
            setSpellTokens(
                ignore = arrayOf(
                    "Ишуи",
                    "шаблонизаторах",
                    "препроцессинг",
                    "десериализации",
                    "постпроцессинга",
                    "сниппеты",
                    "переиспользует",
                    "переиспользуемое",
                    "портируемый",
                    "портируемое",
                    "типографики",
                    "типографику"
                )
            )
            ruleExceptions = setOf(
                "UPPERCASE_SENTENCE_START",
                "RU_UNPAIRED_BRACKETS"
            )
        }
        .check(this)
}

fun saveToPdf(params: String, suffix: String) {
    val options = ChromeOptions()
    options.addArguments("--headless")
    val driver = ChromeDriver(options)
    val url = "file://${project.projectDir}/output/$presentationFile-v$version.html" +
            "?$params"
    println("Converting $url")
    driver[url]
    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30))
    sleep(500)
    val printOptions = PrintOptions().apply {
        pageSize = PageSize(10.8, 19.2)
    }
    val pdf = driver.print(printOptions)
    val pdfContent: ByteArray = Base64.getDecoder().decode(pdf.content)
    Files.write(Paths.get("./output/$presentationFile-${suffix}v$version.pdf"), pdfContent)
}


// tag::defining-custom-element[]
class HorizontalLine() : NoWriterNode() {
    override val isInline: Boolean get() = false
}
// end::defining-custom-element[]

fun odtStyles(): OdtStyleList = OdtStyleList(
    // tag::title-page-styles[]
    OdtStyle { p ->
        if (p !is Paragraph) return@OdtStyle
        if (p.ancestor { it.roles.contains("logo") }.isEmpty()) return@OdtStyle
        attributes("style:master-page-name" to "First_20_Page")
    },
    OdtStyle { tableCell ->
        if (tableCell !is TableCell) return@OdtStyle
        if (tableCell.ancestor { it.roles.contains("about-me") }.isEmpty()) return@OdtStyle
        tableCellProperties {
            arrayOf("top", "right", "bottom", "left")
                .forEach { attributes("fo:border-$it" to "none") }
        }
    },
    // end::title-page-styles[]
    OdtStyle { paragraph ->
        if (paragraph !is Paragraph) return@OdtStyle
        val parent = paragraph.parent() ?: return@OdtStyle
        if (!parent.roles.contains("fio")) return@OdtStyle
        textProperties { attributes("fo:font-weight" to "bold") }
    },
    // tag::custom-element[]
    OdtCustomWriter { horizontalLine ->
        if (horizontalLine !is HorizontalLine) return@OdtCustomWriter
        preOdNode.apply {
            "text:p" {
                attributes("text:style-name" to "Horizontal Line")
                process(horizontalLine)
            }
        }
    },
    // end::custom-element[]
    OdtStyle { p ->
        if (p !is Paragraph) return@OdtStyle
        if (p.ancestor { it.roles.contains("title-photo") }.isEmpty()) return@OdtStyle
        paragraphProperties { attributes("fo:text-align" to "start") }
    },
    // tag::dont-break-slide[]
    OdtStyle { table ->
        if (table !is Table) return@OdtStyle
        tableProperties {
            attributes(
                "fo:keep-with-next" to "always",
                "style:may-break-between-rows" to "false"
            )
        }
    },
    OdtStyle { node ->
        if (node.roles.contains("slide-finish")) return@OdtStyle
        paragraphProperties { attributes("fo:keep-with-next" to "always") }
    }
    // end::dont-break-slide[]
)

fun makeTitle(sourceNode: Node): Node {
    return OpenBlock().apply {
        // tag::extracting-title-element[]
        val title = sourceNode.descendant { it is Heading && it.level == 1 }.first()
        val notes = sourceNode.descendant { it.sourceTagName == "aside" }.first()
        val (fullName, bio, photo, contact, logo) =
            arrayOf("full-name", "bio", "title-photo", "contact", "logo")
                .map { role -> sourceNode.descendant { it.roles.contains(role) }.first() }

        logo.descendant { it is Image }.first().let { it as Image }
            .width = Length(1000F, LengthUnit.cmm)
        photo.descendant { it is Image }.first().let { it as Image }
            .width = Length(1500F, LengthUnit.cmm)
        // end::extracting-title-element[]
        // tag::building-new-title[]
        appendChild(logo)
        appendChild(title)
        table {
            col(Length(18F)); col(Length(152F))
            roles("about-me")
            tableRowGroup(TRG.body) {
                tr {
                    td { appendChild(photo) }
                    td { arrayOf(fullName, bio, contact).forEach { appendChild(it) } }
                }
            }
        }
        appendChild(notes)
        appendChild(Toc(2, "List of slides"))
        normalizeImageDimensions()
        // end::building-new-title[]
    }
}

fun rougeStyles(): OdtStyleList {
    val css = CSSReader.readFromFile(File("${project.projectDir}/syntax.css"))

    val rougeStyles = css?.allStyleRules?.flatMap { styleRule ->
        styleRule.allSelectors.flatMap { selector ->
            styleRule.allDeclarations.map { declaration ->
                selector.allMembers.map { it.asCSSString } to (declaration.property to declaration.expressionAsCSSString)
            }
        }
    }?.filter {
        it.first.size == 3 && it.first[0] == ".highlight" && it.first[2][0] == "."[0] && it.first[2].length <= 3 && arrayOf(
            "color",
            "background-color",
            "font-weight",
            "font-style"
        ).contains(it.second.first)
    }?.map { it.first[2].substring(1) to it.second }?.groupBy { it.first }
        ?.map { it.key to it.value.associate { pair -> pair.second.first to pair.second.second } }?.toMap() ?: mapOf()

    return OdtStyleList(OdtStyle { span ->
        val condition = (span is Span) && (span.ancestor { it is Paragraph && it.sourceTagName == "pre" }.isNotEmpty())
        if (!condition) return@OdtStyle
        rougeStyles.filter { span.roles.contains(it.key) }.forEach { style ->
            textProperties {
                arrayOf("color", "background-color", "font-weight", "font-style").forEach {
                    style.value[it]?.let { value ->
                        attribute("fo:$it", value)
                    }
                }
            }
        }
    })

}

fun unknownTagProcessingRuleRevealJs(): HtmlNode.() -> UnknownTagProcessing = {
    if (classNames().intersect(
            setOf(
                "colist",
                "listingblock",
                "admonitionblock",
                "exampleblock",
                "imageblock",
                "paragraph",
                "book",
                "article",
                "title",
                "footnotes",
                "footnote",
                "dlist",
                "olist",
                "ulist",
                "content",
                *(1..6).map { "sect$it" }.toTypedArray(),
            )
        ).isNotEmpty()
    ) {
        UnknownTagProcessing.PROCESS
    } else if (setOf("dl", "dt", "dd", "section", "aside").contains(nodeName())) {
        UnknownTagProcessing.PROCESS
    } else if (setOf("div", "script", "#data").contains(nodeName())) {
        UnknownTagProcessing.PASS
    } else UnknownTagProcessing.UNDEFINDED
}

