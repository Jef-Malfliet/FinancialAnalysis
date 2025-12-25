package Services;

import Models.DocumentWrapper;
import Models.Enums.ReportStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;
import java.util.function.Function;

public class CompareReportService extends ReportService {

    protected XSSFCellStyle defaultNumberStyle;
    protected XSSFCellStyle greyNumberStyle;
    protected XSSFCellStyle redNumberStyle;
    protected XSSFCellStyle twoDecimalsNumberStyle;
    protected XSSFCellStyle doubleStyle;
    protected XSSFCellStyle percentStyle;
    protected XSSFCellStyle greyStyle;

    public CompareReportService(
            XSSFWorkbook workbook, XSSFSheet reportSheet, XSSFSheet basicValuesSheet,
            List<DocumentWrapper> documents, ReportStyle style, XSSFDataFormat numberFormatter) {
        super(workbook, reportSheet, basicValuesSheet, documents, style, numberFormatter);
    }

    @Override
    protected void createInternal() {
        addCompareReportHeader();

        addCompareReportRow("ACTIVA");

        addCompareReportRow("vlottende activa",
                this::getVlottendeActiva, defaultNumberStyle);

        addCompareReportRow("voorraden en bestellingen in uitvoering",
                this::getVoorradenBestellingenUitvoering, defaultNumberStyle);

        addCompareReportRow("handelsvorderingen", this::getHandelsvorderingen, defaultNumberStyle);

        addCompareReportRow("liquide middelen",
                this::getLiquideMiddelen, defaultNumberStyle);

        addCompareReportRow("totale activa", this::getTotaleActiva, defaultNumberStyle);

        addCompareReportRow("PASSIVA");

        addCompareReportRow("eigen vermogen", this::getEigenVermogen, defaultNumberStyle);

        addCompareReportRow("waarvan reserves",
                this::getReserves, defaultNumberStyle);

        addCompareReportRow("overgedragen winst",
                this::getOverdragenWinstVerlies, defaultNumberStyle);

        addCompareReportRow("totale schulden", this::getTotaleSchulden, defaultNumberStyle);

        addCompareReportRow("schulden op korte termijn", this::getKorteTermijnSchulden, defaultNumberStyle);

        addCompareReportRow("Leveranciersschulden", this::getLeveranciers, defaultNumberStyle);

        addCompareReportRow("RESULTATEN");

        addCompareReportRow("bedrijfsopbrengsten", this::getBedrijfsOpbrengsten, defaultNumberStyle);

        addCompareReportRow("omzet", this::getBedrijfsOpbrengstenOmzet, defaultNumberStyle);

        addCompareReportRow("toegevoegde waarde", this::getToegevoegdeWaarde, defaultNumberStyle);

        addCompareReportRow("brutomarge", this::getBrutoMarge, defaultNumberStyle);

        addCompareReportRow("ebitda", this::getEBITDA, defaultNumberStyle);

        addCompareReportRow("afschrijvingen", this::getAfschrijvingen, defaultNumberStyle);

        addCompareReportRow("bedrijfswinst (ebit)", this::getEBIT, defaultNumberStyle);

        addCompareReportRow("winst van boekjaar na belastingen",
                this::getWinstVerliesBoekjaar, defaultNumberStyle);

        addCompareReportRow("FINANCIËLE RATIO'S");

        addCompareReportRow("cashflow, of kasstroom : nettowinst + afschrijvingen",
                this::getCashFlow, defaultNumberStyle);

        addCompareReportRow("liquiditeitsratio: vlottende activa/schulden op korte termijn (>1)",
                this::getLiquiditeitsRatio, doubleStyle);

        addCompareReportRow("solvabiliteitsratio: schulden op korte termijn/totale activa",
                this::getSolvabiliteitsRatio, percentStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("bedrijfswinst/omzet", i -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("bedrijfswinst/omzet",
                    this::getBedrijfswinstOverOmzet,
                    percentStyle);
        }

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("netto winst/omzet", i -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("netto winst/omzet",
                    this::getNettoWinstOverOmzet, percentStyle);
        }

        addCompareReportRow("rentabiliteitsratio vh eigen vermogen: netto winst/eigen vermogen",
                this::getRentabiliteitsRatioEigenVermogen, percentStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("cash conversion cycle", i -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("cash conversion cycle",
                    this::getCashConversionCycle, defaultNumberStyle);
        }

        addCompareReportRow("netto werkkapitaal (voorraden + vorderingen - leveranciers)",
                this::getNettoWerkkapitaal, defaultNumberStyle);

        skipRow();

        addCompareReportRow("Z score Altman (1,80 > < 2,99)",
                this::getZScoreAltman, CellType.NUMERIC, i -> {

                    ConditionalFormattingRule greenRule = createConditionalFormattingRule(
                            ComparisonOperator.GE, "2.99");

                    PatternFormatting greenPatternFmt = greenRule.createPatternFormatting();
                    greenPatternFmt.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
                    greenPatternFmt.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());

                    ConditionalFormattingRule redRule = createConditionalFormattingRule(
                            ComparisonOperator.LE, "1.80");

                    PatternFormatting redPatternFmt = redRule.createPatternFormatting();
                    redPatternFmt.setFillForegroundColor(IndexedColors.RED.getIndex());
                    redPatternFmt.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());

                    ConditionalFormattingRule greyRule = createConditionalFormattingRule(
                            ComparisonOperator.BETWEEN, "1.80", "2.99");

                    PatternFormatting greyPatternFmt = greyRule.createPatternFormatting();
                    greyPatternFmt.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
                    greyPatternFmt.setFillPattern(FillPatternType.SOLID_FOREGROUND.getCode());

                    CellRangeAddress[] ranges = { CellRangeAddress.valueOf(getCurrentCellAsRange()) };
                    ConditionalFormattingRule[] rules = { greenRule, redRule, greyRule };
                    addConditionalFormatting(ranges, rules);

                    return twoDecimalsNumberStyle;
                }, null, 1);

        addCompareReportRow("PERSONEEL");

        addCompareReportRow("gemiddeld aantal FTE (1003)",
                this::getGemiddeldeAantalFTE, doubleStyle);

        addCompareReportRow("gepresteerde uren (1013)",
                this::getGepresteerdeUren, defaultNumberStyle);

        addCompareReportRow("personeelskosten (1023)",
                this::getPersoneelskosten, defaultNumberStyle);

        skipRow();

        addCompareReportRow("gemiddeld aantal FTE uitzendkrachten (150)",
                this::getGemiddeldeAantalFTEUitzendkrachten, doubleStyle);

        addCompareReportRow("gepresteerde uren uitzendkrachten (151)",
                this::getGepresteerdeUrenUitzendkrachten, defaultNumberStyle);

        addCompareReportRow("personeelskosten uitzendkrachten (152)",
                this::getPersoneelskostenUitzendkrachten, defaultNumberStyle);

        skipRow();

        addCompareReportRow("aantal werknemers op 31/12 (105/3)",
                this::getAantalWerknemersOpEindeBoekjaar, doubleStyle);

        addCompareReportRow("bedienden op 31/12 (134/3)",
                this::getAantalBediendenOpEindeBoekjaar, doubleStyle);

        addCompareReportRow("arbeiders op 31/12 (134/3)",
                this::getAantalArbeidersOpEindeBoekjaar, doubleStyle);

        addCompareReportRow("PERSONEELRATIO'S");

        addCompareReportRow("personeelskost/aantal FTE",
                this::getPersoneelkostOverFTE,
                defaultNumberStyle);

        addCompareReportRow("personeelskost/gepresteerde uren",
                this::getPersoneelskostGepresteerdeUren,
                doubleStyle);

        skipRow();

        addCompareReportRow("personeelskost uitzendkrachten/aantal FTE uitzendkrachten",
                this::getPersoneelkostenUitzendkrachtenOverGemiddeldeFTE,
                defaultNumberStyle);

        addCompareReportRow("personeelskost uitzendkrachten/aantal FTE uitzendkrachten",
                this::getPersoneelkostenUitzendkrachtenOverFTE,
                doubleStyle);

        skipRow();

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("omzet/totaal aantal gepresteerde uren (eigen + interim)",
                    i -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("omzet/totaal aantal gepresteerde uren (eigen + interim)",
                    this::getOmzetOverAantalGepresteerdeUren,
                    doubleStyle, null);
        }

        addCompareReportRow("netto winst/totaal aantal gepresteerde uren (eigen + interim)",
                this::getNettoWinstOverAantalGepresteerdeUren,
                doubleStyle);

        addCompareReportRow("cashflow/totaal aantal gepresteerde uren (eigen + interim)",
                this::getCashflowOverAantalGepresteerdeUren,
                doubleStyle);

        skipRow();

        addCompareReportRow("verhouding arbeiders/bedienden (31/12)",
                this::getVerhoudingArbeidersBedienden,
                doubleStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("omzet/bediende (31/12)",
                    i -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("omzet/bediende (31/12)",
                    this::getOmzetOverAantalBedienden,
                    defaultNumberStyle, null);
        }

        addCompareReportRow("netto winst/bediende (31/12)",
                this::getNettoWinstOverAantalBedienden,
                defaultNumberStyle);

        addCompareReportRow("netto winst bediende (31/12) per uur: netto winst/bediende (31/12)/1744",
                this::getNettoWinstBediendePerUur,
                doubleStyle);
    }

    @Override
    protected void initializeStyles() {
        XSSFFont fontBold = workbook.createFont();
        fontBold.setBold(true);

        greyStyle = workbook.createCellStyle();
        greyStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        greyStyle.setFont(fontBold);
        greyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        defaultNumberStyle = workbook.createCellStyle();
        defaultNumberStyle.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        greyNumberStyle = workbook.createCellStyle();
        greyNumberStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        greyNumberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        greyNumberStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        redNumberStyle = workbook.createCellStyle();
        redNumberStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redNumberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        redNumberStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        twoDecimalsNumberStyle = workbook.createCellStyle();
        twoDecimalsNumberStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        percentStyle = workbook.createCellStyle();
        percentStyle.setDataFormat(numberFormatter.getFormat("### ### ##0.00%"));

        doubleStyle = workbook.createCellStyle();
        doubleStyle.setDataFormat(numberFormatter.getFormat("### ### ##0.00"));
    }

    //<editor-fold desc="addCompareReportRow">

    private void addCompareReportRow(String rowName) {
        addCompareReportRow(rowName, null, null, d -> null, null, 0);
    }

    private void addCompareReportHeader() {
        XSSFRow nameRow = addRow(reportSheet, 2);
        addCell(nameRow, "Naam");

        for (DocumentWrapper document : documents)
            addCell(nameRow, document.getBusiness().getName());

        XSSFRow yearRow = addRow(reportSheet, 2);
        addCell(yearRow, "Boekjaar");

        for (DocumentWrapper document : documents)
            addCell(nameRow, document.getYear());
    }

    private void addCompareReportRow(String rowName, Function<Integer, String> formulaFunction, XSSFCellStyle valueStyle) {
        addCompareReportRow(rowName, formulaFunction, CellType.NUMERIC, d -> valueStyle, null, 1);
    }

    private void addCompareReportRow(String rowName, Function<Integer, String> formulaFunction, XSSFCellStyle valueStyle, XSSFCellStyle titleStyle) {
        addCompareReportRow(rowName, formulaFunction, CellType.NUMERIC, d -> valueStyle, titleStyle, 1);
    }

    private void addCompareReportRow(String rowName, Function<Integer, String> formulaFunction, CellType cellType, Function<DocumentWrapper, XSSFCellStyle> valueStyleFunction, XSSFCellStyle titleStyle, int initialColumnIndex) {
        XSSFRow row = addRow(reportSheet, initialColumnIndex);
        Cell titleCell = row.createCell(initialColumnIndex);

        titleCell.setCellValue(rowName);

        if (titleStyle != null) titleCell.setCellStyle(titleStyle);

        if (formulaFunction != null) {
            int columnIndex = initialColumnIndex + 1;

            for (int i = 0; i < documentCount; i++) {
                if (reportStyle != null) {
                    XSSFCellStyle style = valueStyleFunction.apply(documents.get(i));
                    addReferenceCell(row, columnIndex + i, formulaFunction.apply(i), style, cellType);
                } else {
                    addReferenceCell(row, columnIndex + i, formulaFunction.apply(i), cellType);
                }
            }
        }
    }

    //</editor-fold>

    // <editor-fold desc="getters">

    private String getBedrijfswinstOverOmzet(Integer i) {
        return getBedrijfsWinstVerlies(i).dividedBy(getBedrijfsOpbrengstenOmzet(i));
    }

    private String getPersoneelkostOverFTE(Integer i) {
        return getPersoneelskosten(i).dividedBy(getGemiddeldeAantalFTE(i)).rounded(0);
    }

    private String getPersoneelskostGepresteerdeUren(Integer i) {
        return getPersoneelskosten(i).dividedBy(getGepresteerdeUren(i)).rounded(0);
    }

    private String getPersoneelkostenUitzendkrachtenOverGemiddeldeFTE(Integer i) {
        return getPersoneelskostenUitzendkrachten(i).dividedBy(getGemiddeldeAantalFTEUitzendkrachten(i)).rounded(0);
    }

    private String getPersoneelkostenUitzendkrachtenOverFTE(Integer i) {
        return getPersoneelskostenUitzendkrachten(i).dividedBy(getGepresteerdeUrenUitzendkrachten(i)).rounded(0);
    }

    private String getOmzetOverAantalGepresteerdeUren(Integer i) {
        return getBedrijfsOpbrengstenOmzet(i).dividedBy(getGepresteerdeUrenUitzendkrachten(i)).add(getGepresteerdeUren(i));
    }

    private String getNettoWinstOverAantalGepresteerdeUren(Integer i) {
        return getWinstVerliesBoekjaar(i).dividedBy(getGepresteerdeUrenUitzendkrachten(i)).add(getGepresteerdeUren(i));
    }

    private String getCashflowOverAantalGepresteerdeUren(Integer i) {
        return getCashFlow(i).inParenthesis().dividedBy(getGepresteerdeUrenUitzendkrachten(i).add(getGepresteerdeUren(i)).inParenthesis());
    }

    private String getVerhoudingArbeidersBedienden(Integer i) {
        return getAantalArbeidersOpEindeBoekjaar(i).dividedBy(getAantalBediendenOpEindeBoekjaar(i));
    }

    private String getOmzetOverAantalBedienden(Integer i) {
        return getBedrijfsOpbrengstenOmzet(i).dividedBy(getAantalBediendenOpEindeBoekjaar(i));
    }

    private String getNettoWinstOverAantalBedienden(Integer i) {
        return getWinstVerliesBoekjaar(i).dividedBy(getAantalBediendenOpEindeBoekjaar(i));
    }

    private String getNettoWinstBediendePerUur(Integer i) {
        return getWinstVerliesBoekjaar(i).dividedBy(getAantalBediendenOpEindeBoekjaar(i)).inParenthesis().dividedBy("1744");
    }

    // </editor-fold>
}
