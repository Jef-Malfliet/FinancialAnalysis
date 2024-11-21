package Services;

import Models.DocumentWrapper;
import Models.Enums.ReportStyle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

public class HistoryReportService extends ReportService {

    private HSSFCellStyle BoldGreenCenter;
    private HSSFCellStyle BoldGreen;
    private HSSFCellStyle BoldGreenNumber;
    private HSSFCellStyle Bold;
    private HSSFCellStyle BoldGrey;
    private HSSFCellStyle BoldYellow;
    private HSSFCellStyle BoldLightGrey;
    private HSSFCellStyle BoldBottom;
    private HSSFCellStyle BoldBottomTop;
    private HSSFCellStyle BoldBlue;
    private HSSFCellStyle BottomNormal;
    private HSSFCellStyle BoldTop;
    private HSSFCellStyle BoldNumber;
    private HSSFCellStyle BoldYellowNumber;
    private HSSFCellStyle BoldLightGreyNumber;
    private HSSFCellStyle BoldBottomNumber;
    private HSSFCellStyle BoldTopNumber;
    private HSSFCellStyle BoldTopDouble;
    private HSSFCellStyle PercentStyleBold;
    private HSSFCellStyle DoubleStyleBold;
    private HSSFCellStyle DoubleStyleBoldTop;
    private HSSFCellStyle PercentStyleBoldBottom;
    private HSSFCellStyle Grey;
    private HSSFCellStyle NumberBottomTop;
    private HSSFCellStyle PercentStyleBoldBottomTop;

    public HistoryReportService(
            HSSFWorkbook workbook, HSSFSheet reportSheet, HSSFSheet ratiosSheet,
            List<DocumentWrapper> documents, ReportStyle style, HSSFDataFormat numberFormatter) {
        super(workbook, reportSheet, ratiosSheet, documents, style, numberFormatter);
    }

    @Override
    protected void createInternal() {
        skipRow();

        addHistoryReportHeaderRow(BoldBottomTop, BoldTop, BoldGrey);

        String text1 = documents.get(0).getBusiness().getName() +
                (reportStyle.equals(ReportStyle.HISTORIEKNV) ? " NV" : " BVBA");

        addHistoryReportRow(
                text1, BoldGreenCenter,
                null, null,
                "CORE NETTO WERK KAPITAAL", Bold,
                "voorraden+vorderingen-leveranciers", null,
                d -> Math.round(getCoreNettoWerkKapitaal(d)), BoldNumber);

        skipRow();

        addHistoryReportRow(
                "CAPITAL EMPLOYED", BoldBottom,
                "netto werkkapitaal+vaste activa", BottomNormal,
                d -> Math.round(getCapitalEmployed(d)), BoldBottomNumber);

        addHistoryReportRow("BALANS ACTIVA", BoldGrey, DocumentWrapper::getYear, BoldGrey);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "EBIT MARGE", BoldTop,
                    "EBIT", BoldBottomTop,
                    d -> getEBIT(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "EBIT MARGE", Grey,
                    "EBIT", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "VASTE ACTIVA", BoldYellow,
                d -> Math.round(getVasteActiva(d)), BoldYellowNumber,
                null, null,
                "Omzet", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "IMMATERIELE (evt. Goodwill)", Bold,
                d -> Math.round(getImmaterieleVasteActiva(d)), BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "MATERIELE", Bold,
                    d -> Math.round(getMaterieleVasteActiva(d)), BoldNumber,
                    "EBITDA MARGE", Bold,
                    "EBITDA", BoldBottom,
                    d -> getEBITDA(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "MATERIELE", Bold,
                    d -> Math.round(getMaterieleVasteActiva(d)), BoldNumber,
                    "EBITDA MARGE", Grey,
                    "EBITDA", Grey,
                    d -> "/", Grey);
        }

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getFinancieleVasteActiva(d)), BoldNumber,
                    "> 12-15%", Bold,
                    "OMZET", Bold);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getFinancieleVasteActiva(d)), BoldNumber,
                    "> 12-15%", Grey,
                    "OMZET", Grey);
        }

        addHistoryReportRow(
                "VLOTTENDE ACTIVA", BoldYellow,
                d -> Math.round(getVlottendeActiva(d)), BoldYellowNumber);

        addHistoryReportRow(
                "VOORRADEN", Bold,
                d -> Math.round(getVoorradenEnBestellingenInUitvoering(d)), BoldNumber,
                "RENDEMENT EIGEN", Bold,
                "NETTO WINST", BoldBottom,
                d -> getWinstVerliesBoekjaar(d) / getEigenVermogen(d), PercentStyleBold);

        addHistoryReportRow(
                "HANDELSVORDERINGEN", Bold,
                d -> Math.round(getHandelsvorderingen(d)), BoldNumber,
                "VERMOGEN >?", Bold,
                "EIGEN VERMOGEN", Bold);

        addHistoryReportRow(
                "ANDERE", Bold,
                d -> Math.round(getAndereVorderingen(d)), BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "CASH", Bold,
                    d -> Math.round(getCash(d)), BoldNumber,
                    "ROTATIE", Bold,
                    "OMZET", BoldBottom,
                    d -> getBedrijfsOpbrengsten(d) / getCapitalEmployed(d), DoubleStyleBold);
        } else {
            addHistoryReportRow(
                    "CASH", Bold,
                    d -> Math.round(getCash(d)), BoldNumber,
                    "ROTATIE", Grey,
                    "OMZET", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "TOTALE ACTIVA", BoldYellow,
                d -> Math.round(getTotaleActiva(d)), BoldYellowNumber,
                "CAPITAL EMPLOYED", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        skipRow();

        addHistoryReportRow(
                "BALANS PASSIVA", BoldGrey,
                DocumentWrapper::getYear, BoldGrey,
                "RENDEMENT OP DE", Bold,
                "EBIT", BoldBottom,
                d -> getEBIT(d) / getCapitalEmployed(d), PercentStyleBold);

        addHistoryReportRow(
                "INGEZETTE MIDDELEN (ROCE)", Bold,
                "CAPITAL EMPLOYED", Bold,
                null, null);

        addHistoryReportRow(
                "EIGEN VERMOGEN", BoldYellow,
                d -> Math.round(getEigenVermogen(d)), BoldYellowNumber,
                ">WACC(10%?)", BoldBottom,
                null, BottomNormal,
                null, BottomNormal);

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "SCHULDEN", BoldYellow,
                    d -> Math.round(getLangeTermijnFinancieleSchulden(d) + getKorteTermijnSchulden(d) + getProvisies(d)), BoldYellowNumber,
                    "VOORRAADROTATIE", BoldTop,
                    "VOORRADEN X 365", BoldBottomTop,
                    d -> Math.round(getVoorraadrotatie(d)), BoldTopNumber);
        } else {
            addHistoryReportRow(
                    "SCHULDEN", BoldYellow,
                    d -> Math.round(getLangeTermijnFinancieleSchulden(d) + getKorteTermijnSchulden(d) + getProvisies(d)), BoldYellowNumber,
                    "VOORRAADROTATIE", Grey,
                    "VOORRADEN X 365", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "PROVISIES", Bold,
                d -> Math.round(getProvisies(d)), BoldNumber,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "LANGE TERMIJN SCHULDEN", BoldLightGrey,
                d -> Math.round(getLangeTermijnSchulden(d)), BoldLightGreyNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getLangeTermijnFinancieleSchulden(d)), BoldNumber,
                    "KLANTENKREDIET", Bold,
                    "VORDERINGEN X 365", BoldBottom,
                    d -> Math.round(getKlantenKrediet(d)), BoldNumber);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getLangeTermijnFinancieleSchulden(d)), BoldNumber,
                    "KLANTENKREDIET", Grey,
                    "VORDERINGEN X 365", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "ANDERE", Bold,
                d -> Math.round(getLangeTermijnOverigeSchulden(d)), BoldNumber,
                "DSO", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "KORTE TERMIJN SCHULDEN", BoldLightGrey,
                d -> Math.round(getKorteTermijnSchulden(d)), BoldLightGreyNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getKorteTermijnFinancieleSchulden(d)), BoldNumber,
                    "LEVERANCIERSKREDIET", Bold,
                    "LEVERANCIERS X 365", BoldBottom,
                    d -> Math.round(getLeveranciersKrediet(d)), BoldNumber);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    d -> Math.round(getKorteTermijnFinancieleSchulden(d)), BoldNumber,
                    "LEVERANCIERSKREDIET", Grey,
                    "LEVERANCIERS X 365", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "LEVERANCIERS", Bold,
                d -> Math.round(getLeveranciers(d)), BoldNumber,
                "DPO", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "ANDERE", Bold,
                d -> Math.round(getAndereSchuldenKorteTermijn(d)), BoldNumber);

        addHistoryReportRow(
                "TOTALE PASSIVA", BoldYellow,
                d -> Math.round(getTotalePassiva(d)), BoldYellowNumber,
                "CASH FLOW", Bold,
                "NETTO WINST + AFSCHRIJVINGEN", Bold,
                d -> Math.round(getCashFlow(d)), BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "marge / omzet", Bold,
                    d -> getCashFlow(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "marge / omzet", Grey,
                    d -> "/", Grey);
        }

        skipRow();

        addHistoryReportRow(
                "RESULTATENREKENING", BoldGrey,
                DocumentWrapper::getYear, BoldGrey,
                "FREE CASH FLOW", Bold,
                "CASH FLOW - INVESTERINGEN", Bold,
                d -> Math.round(getCashFlow(d) - getInvesteringen(d)), BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "marge / omzet", Bold,
                    d -> (getCashFlow(d) - getInvesteringen(d)) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "marge / omzet", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "BEDRIJFSOPBRENGSTEN", BoldYellow,
                reportStyle.equals(ReportStyle.HISTORIEKBVBA) ? d -> "/" : d -> Math.round(getBedrijfsOpbrengsten(d)), BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "OMZET", Bold,
                this::getBedrijfsOpbrengstenOmzet, BoldNumber);

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA)) {
            addHistoryReportRow(
                    "BEDRIJFSKOSTEN", BoldYellow,
                    d -> "/", BoldYellowNumber,
                    "CASH CYCLE", BoldBlue,
                    null, BoldBlue,
                    d -> Math.round(getVoorraadrotatie(d) + getKlantenKrediet(d) - getLeveranciersKrediet(d)), BoldBlue);
        } else {
            addHistoryReportRow(
                    "BEDRIJFSKOSTEN", BoldYellow,
                    d -> -Math.round(getTotaleBedrijfskosten(d)), BoldYellowNumber,
                    "CASH CYCLE", Grey,
                    null, Grey,
                    d -> "/", Grey);
        }

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA)) {
            addHistoryReportRow(
                    "AANKOPEN", Grey,
                    d -> "/", Grey,
                    "CURRENT RATIO", BoldTop,
                    "COURANTE ACTIVA", BoldBottomTop,
                    d -> getVlottendeActiva(d) / getKorteTermijnSchulden(d), DoubleStyleBoldTop);
        } else {
            addHistoryReportRow(
                    "AANKOPEN", Bold,
                    d -> Math.round(getAankopen(d)), BoldNumber,
                    "CURRENT RATIO", BoldTop,
                    "COURANTE ACTIVA", BoldBottomTop,
                    d -> getVlottendeActiva(d) / getKorteTermijnSchulden(d), DoubleStyleBoldTop);
        }

        addHistoryReportRow(
                "TOEGEVOEGDE WAARDE", BoldGreen,
                d -> Math.round(getToegevoegdeWaarde(d)), BoldGreenNumber,
                ">1", Bold,
                "COURANTE PASSIVA", Bold);

        skipRow();

        addHistoryReportRow(
                "DIENSTEN EN DIVERSE GOEDEREN", Bold,
                d -> -Math.round(getDienstenEnDiverseGoederen(d)), BoldNumber);

        addHistoryReportRow(
                "BRUTOMARGE", BoldGreen,
                d -> Math.round(getBrutoMarge(d)), BoldGreenNumber,
                "QUICK RATIO", Bold,
                "COURANTE ACTIVA - VOORRAAD", BoldBottom,
                d -> (getVlottendeActiva(d) - getVoorradenEnBestellingenInUitvoering(d)) / getKorteTermijnSchulden(d), DoubleStyleBold);

        addHistoryReportRow(
                ">0,7", Bold,
                "COURANTE PASSIVA", Bold,
                null, null);

        addHistoryReportRow(
                "PERSONEELSKOSTEN", Bold,
                d -> -Math.round(getPersoneelskosten(d)), BoldNumber);

        addHistoryReportRow(
                "ANDERE KOSTEN", Bold,
                d -> -Math.round(getAndereKosten(d) - getDienstenEnDiverseGoederen(d)), BoldNumber);

        addHistoryReportRow(
                "EBITDA", BoldGreen,
                d -> Math.round(getEBITDA(d)), BoldGreenNumber,
                "SOLVABILITEIT", Bold,
                "EIGEN VERMOGEN", BoldBottom,
                d -> getEigenVermogen(d) / getTotalePassiva(d), PercentStyleBold);

        addHistoryReportRow(
                ">25%", Bold,
                "TOTALE PASSIVA", Bold,
                null, null);

        addHistoryReportRow(
                "AFSCHRIJVINGEN", Bold,
                d -> -Math.round(getAfschrijvingen(d)), BoldNumber);

        addHistoryReportRow(
                "WAARDEVERMINDERINGEN", Bold,
                d -> -Math.round(getWaardeVermindering(d)), BoldNumber,
                "GEARING", Bold,
                "NETTO FINANCIELE SCHULDEN", BoldBottom,
                d -> (getKorteTermijnFinancieleSchulden(d) + getLangeTermijnFinancieleSchulden(d) - getCash(d))
                        / getEigenVermogen(d), DoubleStyleBold);

        addHistoryReportRow(
                "<1", BoldBottom,
                "EIGEN VERMOGEN", BoldBottom,
                null, BottomNormal);

        addHistoryReportRow(
                "BEDRIJFSWINST(EBIT)", BoldYellow,
                d -> Math.round(getEBIT(d)), BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "FINANCIELE RESULTATEN", Bold,
                d -> Math.round(getFinancieleResultaten(d)), BoldNumber,
                "FINANCIELE LASTEN:", BoldBottomTop,
                null, BoldBottomTop,
                d -> Math.round(getFinancieleKosten(d)), BoldLightGreyNumber);

        addHistoryReportRow(
                "UITZONDERLIJKE RESULTATEN", Bold,
                d -> Math.round(getUitzonderlijkeResultaten(d)), BoldNumber);

        addHistoryReportRow(
                "FINANCIERINGSLAST", BoldTop,
                "SCHULDEN", BoldBottomTop,
                this::getFinancieringsLast, BoldTopDouble);

        addHistoryReportRow(
                "RESULTAAT VOOR BELASTINGEN", BoldYellow,
                d -> Math.round(getResultaatVoorBelastingen(d)), BoldYellowNumber,
                "<4", Bold,
                "EBITDA", Bold);

        skipRow();

        addHistoryReportRow(
                "BELASTINGEN", Bold,
                d -> -Math.round(getBelastingen(d)), BoldNumber,
                "INTRESTDEKKING", Bold,
                "EBITDA", BoldBottom,
                d -> Math.round(getEBITDA(d) / getFinancieleKosten(d)), BoldNumber);

        addHistoryReportRow(
                ">1", BoldBottom,
                "FINANCIELE LASTEN", BoldBottom,
                null, BottomNormal);

        addHistoryReportRow(
                "NETTO WINST", BoldYellow,
                d -> Math.round(getWinstVerliesBoekjaar(d)), BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "Kostenstructuur", BoldTop,
                null, BoldTop,
                d -> null, BoldTop);

        addHistoryReportRow(
                "EXTRA INFO", Bold,
                null, null,
                null, null,
                null, null,
                null, null);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "Aankopen/omzet", Bold,
                    "AK/omzet", Bold,
                    d -> getAankopen(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Aankopen/omzet", Grey,
                    "AK/omzet", Grey,
                    d -> "/", Grey);
        }

        addHistoryReportRow(
                "Investeringen", BoldBottomTop,
                d -> Math.round(getInvesteringen(d)), NumberBottomTop);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "Investeringen/omzet", BoldBottomTop,
                    d -> getInvesteringen(d) / getBedrijfsOpbrengsten(d), PercentStyleBoldBottomTop,
                    "Personeelskosten/omzet", Bold,
                    "PK/omzet", Bold,
                    d -> getPersoneelskosten(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Investeringen/omzet", Grey,
                    d -> "/", Grey,
                    "Personeelskosten/omzet", Grey,
                    "PK/omzet", Grey,
                    d -> "/", Grey);

        }

        addHistoryReportRow(
                "Investeringen/brutomarge", BoldBottomTop,
                d -> getInvesteringen(d) / getBrutoMarge(d), PercentStyleBoldBottomTop);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "Andere kosten/omzet", Bold,
                    "andere/omzet", Bold,
                    d -> getAndereKosten(d) / getBedrijfsOpbrengsten(d), PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Andere kosten/omzet", Grey,
                    "andere/omzet", Grey,
                    d -> "/", Grey);
        }

        skipRow();

        addHistoryReportRow(
                "Personeelskosten/brutomarge", Bold,
                "PK/brutomarge", Bold,
                d -> getPersoneelskosten(d) / getBrutoMarge(d), PercentStyleBold);

        skipRow();

        addHistoryReportRow(
                "Andere kosten/brutomarge", BoldBottom,
                "AK/brutomarge", BoldBottom,
                d -> getAndereKosten(d) / getBrutoMarge(d), PercentStyleBoldBottom);
    }

    @Override
    protected void initializeStyles() {
        HSSFFont fontBold = workbook.createFont();
        fontBold.setBold(true);

        Bold = workbook.createCellStyle();
        Bold.setFont(fontBold);

        BoldNumber = workbook.createCellStyle();
        BoldNumber.setFont(fontBold);
        BoldNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldGreenCenter = workbook.createCellStyle();
        BoldGreenCenter.setFont(fontBold);
        BoldGreenCenter.setAlignment(HorizontalAlignment.CENTER);
        BoldGreenCenter.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        BoldGreenCenter.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldGreenCenter.setBorderBottom(BorderStyle.THIN);
        BoldGreenCenter.setBorderTop(BorderStyle.THIN);

        BoldGreen = workbook.createCellStyle();
        BoldGreen.setFont(fontBold);
        BoldGreen.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
        BoldGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldGreen.setBorderBottom(BorderStyle.THIN);
        BoldGreen.setBorderTop(BorderStyle.THIN);

        BoldGreenNumber = workbook.createCellStyle();
        BoldGreenNumber.setFont(fontBold);
        BoldGreenNumber.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
        BoldGreenNumber.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldGreenNumber.setBorderBottom(BorderStyle.THIN);
        BoldGreenNumber.setBorderTop(BorderStyle.THIN);
        BoldGreenNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldGrey = workbook.createCellStyle();
        BoldGrey.setFont(fontBold);
        BoldGrey.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        BoldGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldGrey.setBorderBottom(BorderStyle.THIN);
        BoldGrey.setBorderTop(BorderStyle.THIN);

        BoldYellow = workbook.createCellStyle();
        BoldYellow.setFont(fontBold);
        BoldYellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        BoldYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldYellow.setBorderBottom(BorderStyle.THIN);
        BoldYellow.setBorderTop(BorderStyle.THIN);

        BoldYellowNumber = workbook.createCellStyle();
        BoldYellowNumber.setFont(fontBold);
        BoldYellowNumber.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        BoldYellowNumber.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldYellowNumber.setBorderBottom(BorderStyle.THIN);
        BoldYellowNumber.setBorderTop(BorderStyle.THIN);
        BoldYellowNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldLightGrey = workbook.createCellStyle();
        BoldLightGrey.setFont(fontBold);
        BoldLightGrey.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        BoldLightGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldLightGrey.setBorderBottom(BorderStyle.THIN);
        BoldLightGrey.setBorderTop(BorderStyle.THIN);

        BoldLightGreyNumber = workbook.createCellStyle();
        BoldLightGreyNumber.setFont(fontBold);
        BoldLightGreyNumber.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        BoldLightGreyNumber.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldLightGreyNumber.setBorderBottom(BorderStyle.THIN);
        BoldLightGreyNumber.setBorderTop(BorderStyle.THIN);
        BoldLightGreyNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldBlue = workbook.createCellStyle();
        BoldBlue.setFont(fontBold);
        BoldBlue.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        BoldBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldBlue.setBorderBottom(BorderStyle.THIN);
        BoldBlue.setBorderTop(BorderStyle.THIN);

        BoldBottom = workbook.createCellStyle();
        BoldBottom.setFont(fontBold);
        BoldBottom.setBorderBottom(BorderStyle.THIN);

        BoldBottomNumber = workbook.createCellStyle();
        BoldBottomNumber.setFont(fontBold);
        BoldBottomNumber.setBorderBottom(BorderStyle.THIN);
        BoldBottomNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldTop = workbook.createCellStyle();
        BoldTop.setFont(fontBold);
        BoldTop.setBorderTop(BorderStyle.THIN);

        BoldTopNumber = workbook.createCellStyle();
        BoldTopNumber.setFont(fontBold);
        BoldTopNumber.setBorderTop(BorderStyle.THIN);
        BoldTopNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        BoldBottomTop = workbook.createCellStyle();
        BoldBottomTop.setFont(fontBold);
        BoldBottomTop.setBorderBottom(BorderStyle.THIN);
        BoldBottomTop.setBorderTop(BorderStyle.THIN);

        BottomNormal = workbook.createCellStyle();
        BottomNormal.setBorderBottom(BorderStyle.THIN);

        BoldTopDouble = workbook.createCellStyle();
        BoldTopDouble.setFont(fontBold);
        BoldTopDouble.setBorderTop(BorderStyle.THIN);
        BoldTopDouble.setDataFormat(numberFormatter.getFormat("### ### ##0.00"));

        Grey = workbook.createCellStyle();
        Grey.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        Grey.setFont(fontBold);
        Grey.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        PercentStyleBold = workbook.createCellStyle();
        PercentStyleBold.setDataFormat(numberFormatter.getFormat("### ### ##0.00%"));
        PercentStyleBold.setFont(fontBold);

        PercentStyleBoldBottom = workbook.createCellStyle();
        PercentStyleBoldBottom.setDataFormat(numberFormatter.getFormat("### ### ##0.00%"));
        PercentStyleBoldBottom.setFont(fontBold);
        PercentStyleBoldBottom.setBorderBottom(BorderStyle.THIN);

        DoubleStyleBold = workbook.createCellStyle();
        DoubleStyleBold.setDataFormat(numberFormatter.getFormat("### ### ##0.00"));
        DoubleStyleBold.setFont(fontBold);

        DoubleStyleBoldTop = workbook.createCellStyle();
        DoubleStyleBoldTop.setDataFormat(numberFormatter.getFormat("### ### ##0.00"));
        DoubleStyleBoldTop.setFont(fontBold);
        DoubleStyleBoldTop.setBorderTop(BorderStyle.THIN);

        NumberBottomTop = workbook.createCellStyle();
        NumberBottomTop.setFont(fontBold);
        NumberBottomTop.setBorderTop(BorderStyle.THIN);
        NumberBottomTop.setBorderBottom(BorderStyle.THIN);
        NumberBottomTop.setDataFormat(numberFormatter.getFormat("### ### ##0"));

        PercentStyleBoldBottomTop = workbook.createCellStyle();
        PercentStyleBoldBottomTop.setDataFormat(numberFormatter.getFormat("### ### ##0.00%"));
        PercentStyleBoldBottomTop.setFont(fontBold);
        PercentStyleBoldBottomTop.setBorderBottom(BorderStyle.THIN);
        PercentStyleBoldBottomTop.setBorderTop(BorderStyle.THIN);
    }
}
