package Services;

import Models.DocumentWrapper;
import Models.Enums.ReportStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;
import java.util.function.Function;

public class HistoryReportService extends ReportService {

    private XSSFCellStyle BoldGreenCenter;
    private XSSFCellStyle BoldGreen;
    private XSSFCellStyle BoldGreenNumber;
    private XSSFCellStyle Bold;
    private XSSFCellStyle BoldGrey;
    private XSSFCellStyle BoldYellow;
    private XSSFCellStyle BoldLightGrey;
    private XSSFCellStyle BoldBottom;
    private XSSFCellStyle BoldBottomTop;
    private XSSFCellStyle BoldBlue;
    private XSSFCellStyle BoldBlueNumber;
    private XSSFCellStyle BottomNormal;
    private XSSFCellStyle BoldTop;
    private XSSFCellStyle BoldNumber;
    private XSSFCellStyle BoldYellowNumber;
    private XSSFCellStyle BoldLightGreyNumber;
    private XSSFCellStyle BoldBottomNumber;
    private XSSFCellStyle BoldTopNumber;
    private XSSFCellStyle BoldTopDouble;
    private XSSFCellStyle PercentStyleBold;
    private XSSFCellStyle DoubleStyleBold;
    private XSSFCellStyle DoubleStyleBoldTop;
    private XSSFCellStyle PercentStyleBoldBottom;
    private XSSFCellStyle Grey;
    private XSSFCellStyle NumberBottomTop;
    private XSSFCellStyle PercentStyleBoldBottomTop;

    public HistoryReportService(
            XSSFWorkbook workbook, XSSFSheet reportSheet, XSSFSheet ratiosSheet,
            List<DocumentWrapper> documents, ReportStyle style, XSSFDataFormat numberFormatter) {
        super(workbook, reportSheet, ratiosSheet, documents, style, numberFormatter);
    }

    @Override
    protected void createInternal() {
        skipRow();

        addHistoryReportHeaderRow(BoldBottomTop, BoldTop, BoldGrey);

        String title = documents.get(0).getBusiness().getName() +
                (reportStyle.equals(ReportStyle.HISTORIEKNV) ? " NV" : " BVBA");

        addHistoryReportRow(
                title, BoldGreenCenter,
                null, null,
                "CORE NETTO WERK KAPITAAL", Bold,
                "voorraden+vorderingen-leveranciers", null,
                this::getCoreNettoWerkKapitaal, BoldNumber);

        skipRow();

        addHistoryReportRow(
                "CAPITAL EMPLOYED", BoldBottom,
                "netto werkkapitaal+vaste activa", BottomNormal,
                this::getCapitalEmployed, BoldBottomNumber);

        addHistoryReportRow("BALANS ACTIVA", BoldGrey, this::getYear, BoldGrey);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "EBIT MARGE", BoldTop,
                    "EBIT", BoldBottomTop,
                    this::getEBITMarge, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "EBIT MARGE", Grey,
                    "EBIT", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "VASTE ACTIVA", BoldYellow,
                this::getVasteActiva, BoldYellowNumber,
                null, null,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "IMMATERIELE (evt. Goodwill)", Bold,
                this::getImmaterieleVasteActiva, BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "MATERIELE", Bold,
                    this::getMaterieleVasteActiva, BoldNumber,
                    "EBITDA MARGE", Bold,
                    "EBITDA", BoldBottom,
                    this::getEBITDAMarge, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "MATERIELE", Bold,
                    this::getMaterieleVasteActiva, BoldNumber,
                    "EBITDA MARGE", Grey,
                    "EBITDA", Grey,
                    i -> "/", Grey);
        }

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getFinancieleVasteActiva, BoldNumber,
                    "> 12-15%", Bold,
                    "OMZET", Bold);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getFinancieleVasteActiva, BoldNumber,
                    "> 12-15%", Grey,
                    "OMZET", Grey);
        }

        addHistoryReportRow(
                "VLOTTENDE ACTIVA", BoldYellow,
                this::getVlottendeActiva, BoldYellowNumber);

        addHistoryReportRow(
                "VOORRADEN", Bold,
                this::getVoorradenBestellingenUitvoering, BoldNumber,
                "RENDEMENT EIGEN", Bold,
                "NETTO WINST", BoldBottom,
                this::getRendementEigenVermogen, PercentStyleBold);

        addHistoryReportRow(
                "HANDELSVORDERINGEN", Bold,
                this::getHandelsvorderingen, BoldNumber,
                "VERMOGEN >?", Bold,
                "EIGEN VERMOGEN", Bold);

        addHistoryReportRow(
                "ANDERE", Bold,
                this::getVorderingenHoogstens1JaarOverigeVorderingen, BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "CASH", Bold,
                    this::getCash, BoldNumber,
                    "ROTATIE", Bold,
                    "OMZET", BoldBottom,
                    this::getRotatie, DoubleStyleBold);
        } else {
            addHistoryReportRow(
                    "CASH", Bold,
                    this::getCash, BoldNumber,
                    "ROTATIE", Grey,
                    "OMZET", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "TOTALE ACTIVA", BoldYellow,
                this::getTotaleActiva, BoldYellowNumber,
                "", null,
                "CAPITAL EMPLOYED", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        skipRow();

        addHistoryReportRow(
                "BALANS PASSIVA", BoldGrey,
                this::getYear, BoldGrey,
                "RENDEMENT OP DE", Bold,
                "EBIT", BoldBottom,
                this::getRendementOpIngezetteMiddelen, PercentStyleBold);

        addHistoryReportRow(
                "INGEZETTE MIDDELEN (ROCE)", Bold,
                "CAPITAL EMPLOYED", Bold,
                null, null);

        addHistoryReportRow(
                "EIGEN VERMOGEN", BoldYellow,
                this::getEigenVermogen, BoldYellowNumber,
                ">WACC(10%?)", BoldBottom,
                null, BottomNormal,
                i -> null, BottomNormal);

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "SCHULDEN", BoldYellow,
                    this::getSchulden, BoldYellowNumber,
                    "VOORRAADROTATIE", BoldTop,
                    "VOORRADEN X 365", BoldBottomTop,
                    this::getVoorraadrotatie, BoldTopNumber);
        } else {
            addHistoryReportRow(
                    "SCHULDEN", BoldYellow,
                    this::getSchulden, BoldYellowNumber,
                    "VOORRAADROTATIE", Grey,
                    "VOORRADEN X 365", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "PROVISIES", Bold,
                this::getProvisies, BoldNumber,
                "", null,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "LANGE TERMIJN SCHULDEN", BoldLightGrey,
                this::getLangeTermijnSchulden, BoldLightGreyNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getLangeTermijnFinancieleSchulden, BoldNumber,
                    "KLANTENKREDIET", Bold,
                    "VORDERINGEN X 365", BoldBottom,
                    this::getKlantenKrediet, BoldNumber);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getLangeTermijnFinancieleSchulden, BoldNumber,
                    "KLANTENKREDIET", Grey,
                    "VORDERINGEN X 365", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "ANDERE", Bold,
                this::getLangeTermijnOverigeSchulden, BoldNumber,
                "DSO", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "KORTE TERMIJN SCHULDEN", BoldLightGrey,
                this::getKorteTermijnSchulden, BoldLightGreyNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getKorteTermijnFinancieleSchulden, BoldNumber,
                    "LEVERANCIERSKREDIET", Bold,
                    "LEVERANCIERS X 365", BoldBottom,
                    this::getLeveranciersKrediet, BoldNumber);
        } else {
            addHistoryReportRow(
                    "FINANCIELE", Bold,
                    this::getKorteTermijnFinancieleSchulden, BoldNumber,
                    "LEVERANCIERSKREDIET", Grey,
                    "LEVERANCIERS X 365", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "LEVERANCIERS", Bold,
                this::getLeveranciers, BoldNumber,
                "DPO", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey,
                "OMZET", reportStyle.equals(ReportStyle.HISTORIEKNV) ? Bold : Grey);

        addHistoryReportRow(
                "ANDERE", Bold,
                this::getAndereSchuldenKorteTermijn, BoldNumber);

        addHistoryReportRow(
                "TOTALE PASSIVA", BoldYellow,
                this::getTotalePassiva, BoldYellowNumber,
                "CASH FLOW", Bold,
                "NETTO WINST + AFSCHRIJVINGEN", Bold,
                this::getCashFlow, BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "", null,
                    null, null,
                    "marge / omzet", Bold,
                    "", null,
                    this::getCashFlowMargeOverOmzet, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "", null,
                    null, null,
                    "marge / omzet", Grey,
                    "", null,
                    i -> "/", Grey);
        }

        skipRow();

        addHistoryReportRow(
                "RESULTATENREKENING", BoldGrey,
                this::getYear, BoldGrey,
                "FREE CASH FLOW", Bold,
                "CASH FLOW - INVESTERINGEN", Bold,
                this::getFreeCashFlow, BoldNumber);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "", null,
                    null, null,
                    "marge / omzet", Bold,
                    "", null,
                    this::getFreeCashFlowMargeOverOmzet, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "", null,
                    null, null,
                    "marge / omzet", Grey,
                    "", null,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "BEDRIJFSOPBRENGSTEN", BoldYellow,
                reportStyle.equals(ReportStyle.HISTORIEKBVBA) ? i -> "/" : this::getBedrijfsOpbrengsten, BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "OMZET", Bold,
                this::getBedrijfsOpbrengstenOmzet, BoldNumber);

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "BEDRIJFSKOSTEN", BoldYellow,
                    i -> "/", BoldYellowNumber,
                    "CASH CYCLE", BoldBlue,
                    null, BoldBlue,
                    this::getCashCycle, BoldBlueNumber);
        } else {
            addHistoryReportRow(
                    "BEDRIJFSKOSTEN", BoldYellow,
                    i -> getTotaleBedrijfskosten(i).negate(), BoldYellowNumber,
                    "CASH CYCLE", Grey,
                    null, Grey,
                    i -> "/", Grey);
        }

        skipRow();

        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA)) {
            addHistoryReportRow(
                    "AANKOPEN", Grey,
                    i -> "/", Grey,
                    "CURRENT RATIO", BoldTop,
                    "COURANTE ACTIVA", BoldBottomTop,
                    this::getLiquiditeitsRatio, DoubleStyleBoldTop);
        } else {
            addHistoryReportRow(
                    "AANKOPEN", Bold,
                    i -> getAankopen(i).negate(), BoldNumber,
                    "CURRENT RATIO", BoldTop,
                    "COURANTE ACTIVA", BoldBottomTop,
                    this::getLiquiditeitsRatio, DoubleStyleBoldTop);
        }

        addHistoryReportRow(
                "TOEGEVOEGDE WAARDE", BoldGreen,
                this::getToegevoegdeWaarde, BoldGreenNumber,
                ">1", Bold,
                "COURANTE PASSIVA", Bold);

        skipRow();

        addHistoryReportRow(
                "DIENSTEN EN DIVERSE GOEDEREN", Bold,
                i -> getDienstenEnDiverseGoederen(i).negate(), BoldNumber);

        addHistoryReportRow(
                "BRUTOMARGE", BoldGreen,
                this::getBrutoMarge, BoldGreenNumber,
                "QUICK RATIO", Bold,
                "COURANTE ACTIVA - VOORRAAD", BoldBottom,
                this::getCouranteActiveZonderVoorraad, DoubleStyleBold);

        addHistoryReportRow(
                ">0,7", Bold,
                "COURANTE PASSIVA", Bold,
                null, null);

        addHistoryReportRow(
                "PERSONEELSKOSTEN", Bold,
                i -> getPersoneelskosten(i).negate(), BoldNumber);

        addHistoryReportRow(
                "ANDERE KOSTEN", Bold,
                this::getAndereKostenZonderDienstenEnDiverseGoederen, BoldNumber);

        addHistoryReportRow(
                "EBITDA", BoldGreen,
                this::getEBITDA, BoldGreenNumber,
                "SOLVABILITEIT", Bold,
                "EIGEN VERMOGEN", BoldBottom,
                this::getSolvabiliteit, PercentStyleBold);

        addHistoryReportRow(
                ">25%", Bold,
                "TOTALE PASSIVA", Bold,
                null, null);

        addHistoryReportRow(
                "AFSCHRIJVINGEN", Bold,
                i -> getAfschrijvingen(i).negate(), BoldNumber);

        addHistoryReportRow(
                "WAARDEVERMINDERINGEN", Bold,
                i -> getWaardeVermindering(i).negate(), BoldNumber,
                "GEARING", Bold,
                "NETTO FINANCIELE SCHULDEN", BoldBottom,
                this::getGearing, DoubleStyleBold);

        addHistoryReportRow(
                "<1", BoldBottom,
                "EIGEN VERMOGEN", BoldBottom,
                i -> null, BottomNormal);

        addHistoryReportRow(
                "BEDRIJFSWINST(EBIT)", BoldYellow,
                this::getEBIT, BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "FINANCIELE RESULTATEN", Bold,
                this::getFinancieleResultaten, BoldNumber,
                "FINANCIELE LASTEN:", BoldBottomTop,
                null, BoldBottomTop,
                this::getFinancieleKosten, BoldLightGreyNumber);

        addHistoryReportRow(
                "UITZONDERLIJKE RESULTATEN", Bold,
                this::getUitzonderlijkeResultaten, BoldNumber);

        addHistoryReportRow(
                "FINANCIERINGSLAST", BoldTop,
                "SCHULDEN", BoldBottomTop,
                this::getFinancieringsLast, BoldTopDouble);

        addHistoryReportRow(
                "RESULTAAT VOOR BELASTINGEN", BoldYellow,
                this::getResultaatVoorBelastingen, BoldYellowNumber,
                "<4", Bold,
                "EBITDA", Bold);

        skipRow();

        addHistoryReportRow(
                "BELASTINGEN", Bold,
                i -> getBelastingen(i).inParenthesis().negate(), BoldNumber,
                "INTRESTDEKKING", Bold,
                "EBITDA", BoldBottom,
                this::getIntrestdekking, BoldNumber);

        addHistoryReportRow(
                ">1", BoldBottom,
                "FINANCIELE LASTEN", BoldBottom,
                i -> null, BottomNormal);

        addHistoryReportRow(
                "NETTO WINST", BoldYellow,
                this::getWinstVerliesBoekjaar, BoldYellowNumber);

        skipRow();

        addHistoryReportRow(
                "Kostenstructuur", BoldTop,
                null, BoldTop,
                i -> null, BoldTop);

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
                    this::getAankopenOverOmzet, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Aankopen/omzet", Grey,
                    "AK/omzet", Grey,
                    i -> "/", Grey);
        }

        addHistoryReportRow(
                "Investeringen", BoldBottomTop,
                this::getInvesteringen, NumberBottomTop);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "Investeringen/omzet", BoldBottomTop,
                    this::getInvesteringenOverOmzet, PercentStyleBoldBottomTop,
                    "Personeelskosten/omzet", Bold,
                    "PK/omzet", Bold,
                    this::getPersoneelskostenOverOmzet, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Investeringen/omzet", Grey,
                    i -> "/", Grey,
                    "Personeelskosten/omzet", Grey,
                    "PK/omzet", Grey,
                    i -> "/", Grey);

        }

        addHistoryReportRow(
                "Investeringen/brutomarge", BoldBottomTop,
                this::getInvesteringenOverBrutomarge, PercentStyleBoldBottomTop);

        if (reportStyle.equals(ReportStyle.HISTORIEKNV)) {
            addHistoryReportRow(
                    "Andere kosten/omzet", Bold,
                    "andere/omzet", Bold,
                    this::getAndereKostenOverOmzet, PercentStyleBold);
        } else {
            addHistoryReportRow(
                    "Andere kosten/omzet", Grey,
                    "andere/omzet", Grey,
                    i -> "/", Grey);
        }

        skipRow();

        addHistoryReportRow(
                "Personeelskosten/brutomarge", Bold,
                "PK/brutomarge", Bold,
                this::getPersoneelskostenOverBrutomarge, PercentStyleBold);

        skipRow();

        addHistoryReportRow(
                "Andere kosten/brutomarge", BoldBottom,
                "AK/brutomarge", BoldBottom,
                this::getAndereKostenOverBrutomarge, PercentStyleBoldBottom);
    }

    @Override
    protected void initializeStyles() {
        XSSFFont fontBold = workbook.createFont();
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

        BoldBlueNumber = workbook.createCellStyle();
        BoldBlueNumber.setFont(fontBold);
        BoldBlueNumber.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        BoldBlueNumber.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        BoldBlueNumber.setBorderBottom(BorderStyle.THIN);
        BoldBlueNumber.setBorderTop(BorderStyle.THIN);
        BoldBlueNumber.setDataFormat(numberFormatter.getFormat("### ### ##0"));

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

    //<editor-fold desc="addHistoryReportRow">

    private void addHistoryReportHeaderRow(XSSFCellStyle titleStyle, XSSFCellStyle documentStyle1, XSSFCellStyle documentStyle2) {
        XSSFRow row = addRow(reportSheet);

        addCell(row, "NAAM", titleStyle);
        addCell(row, "PER", titleStyle);
        addCell(row, "31/12", titleStyle);

        skipCell(documentCount - 1);

        addCell(row, null, documentStyle1);
        addCell(row, null, documentStyle1);

        for (DocumentWrapper document : documents) {
            addCell(row, document.getYear(), documentStyle2);
        }
    }

    private void addHistoryReportRow(String text1, XSSFCellStyle textStyle1, Function<Integer, String> formulaFunction1, XSSFCellStyle valueStyle1) {
        addHistoryReportRow(text1, textStyle1, formulaFunction1, valueStyle1, null, null, null, null, null, null);
    }

    private void addHistoryReportRow(String text2, XSSFCellStyle textStyle2, String text3, XSSFCellStyle textStyle3, Function<Integer, String> formulaFunction2, XSSFCellStyle valueStyle2) {
        addHistoryReportRow(null, null, null, null, text2, textStyle2, text3, textStyle3, formulaFunction2, valueStyle2);
    }

    private void addHistoryReportRow(String text1, XSSFCellStyle textStyle1, Function<Integer, String> formulaFunction1, XSSFCellStyle valueStyle1, String text2, XSSFCellStyle textStyle2, String text3, XSSFCellStyle textStyle3) {
        addHistoryReportRow(text1, textStyle1, formulaFunction1, valueStyle1, text2, textStyle2, text3, textStyle3, null, null);
    }

    private void addHistoryReportRow(String text1, XSSFCellStyle textStyle1, Function<Integer, String> formulaFunction1, XSSFCellStyle valueStyle1, String text2, XSSFCellStyle textStyle2, String text3, XSSFCellStyle textStyle3, Function<Integer, String> formulaFunction2, XSSFCellStyle valueStyle2) {
        XSSFRow row = addRow(reportSheet);
        addHistoryRowTextCell(row, text1, textStyle1);
        addHistoryRowDocumentValueCell(row, formulaFunction1, valueStyle1);
        skipCell();
        addHistoryRowTextCell(row, text2, textStyle2);
        addHistoryRowTextCell(row, text3, textStyle3);
        addHistoryRowDocumentValueCell(row, formulaFunction2, valueStyle2);
    }

    private void addHistoryRowTextCell(XSSFRow row, String text, XSSFCellStyle cellStyle) {
        if (text != null || cellStyle != null) addCell(row, text, cellStyle);
        else skipCell();
    }

    private void addHistoryRowDocumentValueCell(XSSFRow row, Function<Integer, String> formulaFunction1, XSSFCellStyle cellStyle) {
        if (formulaFunction1 != null || cellStyle != null) {
            for (int i = 0; i < documentCount; i++) {
                addReferenceCell(row, formulaFunction1 == null ? "/" : formulaFunction1.apply(i), cellStyle);
            }
        } else skipCell(documentCount);
    }

    //</editor-fold>

    // <editor-fold desc="getters">

    private String getEBITMarge(int i) {
        return getEBIT(i).inParenthesis().dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getEBITDAMarge(int i) {
        return getEBITDA(i).inParenthesis().dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getRendementEigenVermogen(int i) {
        return getWinstVerliesBoekjaar(i).dividedBy(getEigenVermogen(i));
    }

    private String getRotatie(int i) {
        return getBedrijfsOpbrengsten(i).dividedBy(getCapitalEmployed(i).inParenthesis());
    }

    private String getRendementOpIngezetteMiddelen(int i) {
        return getEBIT(i).inParenthesis().dividedBy(getCapitalEmployed(i).inParenthesis());
    }

    private String getSchulden(int i) {
        return getLangeTermijnFinancieleSchulden(i).add(getKorteTermijnSchulden(i)).add(getProvisies(i));
    }

    private String getCashFlowMargeOverOmzet(int i) {
        return getCashFlow(i).inParenthesis().dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getFreeCashFlow(int i) {
        return getCashFlow(i).subtract(getInvesteringen(i).inParenthesis());
    }

    private String getFreeCashFlowMargeOverOmzet(int i) {
        return getFreeCashFlow(i).dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getCashCycle(int i) {
        return getVoorraadrotatie(i).add(getKlantenKrediet(i)).subtract(getLeveranciersKrediet(i));
    }

    private String getCouranteActiveZonderVoorraad(int i) {
        return getVlottendeActiva(i).subtract(getVoorradenBestellingenUitvoering(i)).inParenthesis().dividedBy(getKorteTermijnSchulden(i).inParenthesis());
    }

    private String getAndereKostenZonderDienstenEnDiverseGoederen(int i) {
        return getAndereKosten(i).inParenthesis().subtract(getDienstenEnDiverseGoederen(i)).inParenthesis().negate();
    }

    private String getSolvabiliteit(int i) {
        return getEigenVermogen(i).dividedBy(getTotalePassiva(i));
    }

    private String getGearing(int i) {
        return getKorteTermijnFinancieleSchulden(i).add(getLangeTermijnFinancieleSchulden(i)).subtract(getCash(i)).inParenthesis().dividedBy(getEigenVermogen(i));
    }

    private String getIntrestdekking(int i) {
        return getEBITDA(i).inParenthesis().dividedBy(getFinancieleKosten(i));
    }

    private String getAankopenOverOmzet(int i) {
        return getAankopen(i).dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getInvesteringenOverOmzet(int i) {
        return getInvesteringen(i).inParenthesis().dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getPersoneelskostenOverOmzet(int i) {
        return getPersoneelskosten(i).dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getAndereKostenOverOmzet(int i) {
        return getAndereKostenZonderDienstenEnDiverseGoederen(i).inParenthesis().dividedBy(getBedrijfsOpbrengsten(i));
    }

    private String getInvesteringenOverBrutomarge(int i) {
        return getInvesteringen(i).inParenthesis().dividedBy(getBrutoMarge(i).inParenthesis());
    }

    private String getPersoneelskostenOverBrutomarge(int i) {
        return getPersoneelskosten(i).dividedBy(getBrutoMarge(i).inParenthesis());
    }

    private String getAndereKostenOverBrutomarge(int i) {
        return getAndereKostenZonderDienstenEnDiverseGoederen(i).inParenthesis().dividedBy(getBrutoMarge(i).inParenthesis());
    }

    // </editor-fold>
}
