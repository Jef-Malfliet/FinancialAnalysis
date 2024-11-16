package Services;

import Models.DocumentWrapper;
import Models.Enums.ReportStyle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

public class CompareReportService extends ReportService {

    protected HSSFCellStyle defaultNumberStyle;
    protected HSSFCellStyle greyNumberStyle;
    protected HSSFCellStyle redNumberStyle;
    protected HSSFCellStyle greenNumberStyle;
    protected HSSFCellStyle doubleStyle;
    protected HSSFCellStyle percentStyle;
    protected HSSFCellStyle greyStyle;

    public CompareReportService(
            HSSFWorkbook workbook, HSSFSheet reportSheet, HSSFSheet basicValuesSheet,
            List<DocumentWrapper> documents, ReportStyle style, HSSFDataFormat numberFormatter) {
        super(workbook, reportSheet, basicValuesSheet, documents, style, numberFormatter);
    }

    @Override
    protected void createInternal() {
        addCompareReportRow("Naam", d -> d.getBusiness().getName());

        addCompareReportRow("Boekjaar", DocumentWrapper::getYear);

        addCompareReportRow("ACTIVA");

        addCompareReportRow("vlottende activa", d -> Math.round(getVlottendeActiva(d)), defaultNumberStyle);

        addCompareReportRow("voorraden en bestellingen in uitvoering", d -> Math.round(getVoorradenEnBestellingenInUitvoering(d)), defaultNumberStyle);

        addCompareReportRow("handelsvorderingen", d -> Math.round(getHandelsvorderingen(d)), defaultNumberStyle);

        addCompareReportRow("liquide middelen", d -> Math.round(getLiquideMiddelen(d)), defaultNumberStyle);

        addCompareReportRow("totale activa", d -> Math.round(getTotaleActiva(d)), defaultNumberStyle);

        addCompareReportRow("PASSIVA");

        addCompareReportRow("eigen vermogen", d -> Math.round(getEigenVermogen(d)), defaultNumberStyle);

        addCompareReportRow("waarvan reserves", d -> Math.round(getReserves(d)), defaultNumberStyle);

        addCompareReportRow("overgedragen winst", d -> Math.round(getOverdragenWinstVerlies(d)), defaultNumberStyle);

        addCompareReportRow("totale schulden", d -> Math.round(getTotaleSchulden(d)), defaultNumberStyle);

        addCompareReportRow("schulden op korte termijn", d -> Math.round(getKorteTermijnSchulden(d)), defaultNumberStyle);

        addCompareReportRow("Leveranciersschulden", d -> Math.round(getLeveranciers(d)), defaultNumberStyle);

        addCompareReportRow("RESULTATEN");

        addCompareReportRow("bedrijfsopbrengsten", d -> Math.round(getBedrijfsOpbrengsten(d)), defaultNumberStyle);

        addCompareReportRow("omzet", d -> Math.round(getBedrijfsOpbrengstenOmzet(d)), defaultNumberStyle);

        addCompareReportRow("toegevoegde waarde", d -> Math.round(getToegevoegdeWaarde(d)), defaultNumberStyle);

        addCompareReportRow("brutomarge", d -> Math.round(getBrutoMarge(d)), defaultNumberStyle);

        addCompareReportRow("ebitda", d -> Math.round(getEBITDA(d)), defaultNumberStyle);

        addCompareReportRow("afschrijvingen", d -> Math.round(getAfschrijvingen(d)), defaultNumberStyle);

        addCompareReportRow("bedrijfswinst (ebit)", d -> Math.round(getEBIT(d)), defaultNumberStyle);
        ;

        addCompareReportRow("winst van boekjaar na belastingen", d -> Math.round(getWinstVerliesBoekjaar(d)), defaultNumberStyle);
        ;

        addCompareReportRow("FINANCIËLE RATIO'S");

        addCompareReportRow("cashflow, of kasstroom : nettowinst + afschrijvingen",
                d -> Math.round(getWinstVerliesBoekjaar(d) + getAfschrijvingen(d)), defaultNumberStyle);

        addCompareReportRow("liquiditeitsratio: vlottende activa/schulden op korte termijn (>1)",
                d -> Math.round(getVlottendeActiva(d) / getKorteTermijnSchulden(d)), doubleStyle);

        addCompareReportRow("solvabiliteitsratio: schulden op korte termijn/totale activa",
                d -> Math.round(getKorteTermijnSchulden(d) / getTotaleActiva(d)), percentStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("bedrijfswinst/omzet", d -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("bedrijfswinst/omzet",
                    d -> Math.round(getBedrijfswinstVerlies(d) / getBedrijfsOpbrengstenOmzet(d)),
                    percentStyle);
        }

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("netto winst/omzet", d -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("netto winst/omzet",
                    d -> Math.round(getWinstVerliesBoekjaar(d) / getBedrijfsOpbrengstenOmzet(d)),
                    percentStyle);
        }

        addCompareReportRow("rentabiliteitsratio vh eigen vermogen: netto winst/eigen vermogen",
                d -> getWinstVerliesBoekjaar(d) / getEigenVermogen(d), percentStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("cash conversion cycle", d -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("cash conversion cycle",
                    d -> Math.round(getVoorraadrotatie(d) + getKlantenKrediet(d) - getLeveranciersKrediet(d)),
                    defaultNumberStyle);
        }

        addCompareReportRow("netto werkkapitaal (voorraden + vorderingen - leveranciers)",
                this::getNettoWerkkapitaal, defaultNumberStyle);

        skipRow();

        addCompareReportRow("Z score Altman (1,80 > < 2,99)",
                this::getZScoreAltman, CellType.NUMERIC, d -> {
                    double value = getZScoreAltman(d);
                    if (value >= 2.99) return greenNumberStyle;
                    if (value <= 1.80) return redNumberStyle;
                    return greyNumberStyle;
                }, null, 1);

        addCompareReportRow("PERSONEEL");

        addCompareReportRow("gemiddeld aantal FTE (1003)",
                this::getGemiddeldeAantalFTE, doubleStyle);

        addCompareReportRow("gepresteerde uren (1013)",
                this::getGepresteerdeUren, defaultNumberStyle);

        addCompareReportRow("personeelskosten (1023)",
                this::getSBPersoneelsKosten, defaultNumberStyle);

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
                this::getAantalArbeiderssOpEindeBoekjaar, doubleStyle);

        addCompareReportRow("PERSONEELRATIO'S");

        addCompareReportRow("personeelskost/aantal FTE",
                d -> Math.round(getSBPersoneelsKosten(d) / getGemiddeldeAantalFTE(d)),
                defaultNumberStyle);

        addCompareReportRow("personeelskost/gepresteerde uren",
                d -> Math.round(getSBPersoneelsKosten(d) / getGepresteerdeUren(d)),
                doubleStyle);

        skipRow();

        addCompareReportRow("personeelskost uitzendkrachten/aantal FTE uitzendkrachten",
                d -> Math.round(getPersoneelskostenUitzendkrachten(d) / getGemiddeldeAantalFTEUitzendkrachten(d)),
                defaultNumberStyle);

        addCompareReportRow("personeelskost uitzendkrachten/aantal FTE uitzendkrachten",
                d -> Math.round(getPersoneelskostenUitzendkrachten(d) / getGepresteerdeUrenUitzendkrachten(d)),
                doubleStyle);

        skipRow();

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("omzet/totaal aantal gepresteerde uren (eigen + interim)",
                    d -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("omzet/totaal aantal gepresteerde uren (eigen + interim)",
                    d -> getBedrijfsOpbrengstenOmzet(d) / (getGepresteerdeUrenUitzendkrachten(d) + getGepresteerdeUren(d)),
                    doubleStyle, null);
        }

        addCompareReportRow("netto winst/totaal aantal gepresteerde uren (eigen + interim)",
                d -> getWinstVerliesBoekjaar(d) / (getGepresteerdeUrenUitzendkrachten(d) + getGepresteerdeUren(d)),
                doubleStyle);

        addCompareReportRow("cashflow/totaal aantal gepresteerde uren (eigen + interim)",
                d -> getCashFlow(d) / (getGepresteerdeUrenUitzendkrachten(d) + getGepresteerdeUren(d)),
                doubleStyle);

        skipRow();

        addCompareReportRow("verhouding arbeiders/bedienden (31/12)",
                d -> getAantalArbeiderssOpEindeBoekjaar(d) / getAantalBediendenOpEindeBoekjaar(d),
                doubleStyle);

        if (reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            addCompareReportRow("omzet/bediende (31/12)",
                    d -> "/", greyStyle, greyStyle);
        } else {
            addCompareReportRow("omzet/bediende (31/12)",
                    d -> getBedrijfsOpbrengstenOmzet(d) / getAantalBediendenOpEindeBoekjaar(d),
                    defaultNumberStyle, null);
        }

        addCompareReportRow("netto winst/bediende (31/12)",
                d -> getWinstVerliesBoekjaar(d) / getAantalBediendenOpEindeBoekjaar(d),
                defaultNumberStyle);

        addCompareReportRow("netto winst bediende (31/12) per uur: netto winst/bediende (31/12)/1744",
                d -> (getWinstVerliesBoekjaar(d) / getAantalBediendenOpEindeBoekjaar(d)) / 1744,
                doubleStyle);
    }

    @Override
    protected void initializeStyles() {
        HSSFFont fontBold = workbook.createFont();
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

        greenNumberStyle = workbook.createCellStyle();
        greenNumberStyle.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
        greenNumberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        greenNumberStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        percentStyle = workbook.createCellStyle();
        percentStyle.setDataFormat(numberFormatter.getFormat("### ### ##0.00%"));

        doubleStyle = workbook.createCellStyle();
        doubleStyle.setDataFormat(numberFormatter.getFormat("### ### ##0.00"));
    }
}
