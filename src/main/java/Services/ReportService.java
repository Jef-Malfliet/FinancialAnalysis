package Services;

import Models.Enums.PropertyName;
import Models.Enums.ReportStyle;
import Models.DocumentWrapper;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.util.List;
import java.util.function.Function;

public abstract class ReportService {

    private final HSSFSheet reportSheet;
    private final HSSFSheet basicValuesSheet;

    protected final List<DocumentWrapper> documents;
    protected final int documentCount;
    protected final ReportStyle reportStyle;
    protected final HSSFWorkbook workbook;
    protected final HSSFDataFormat numberFormatter;
    protected int rowNumber;
    protected int columnNumber;

    public ReportService(
            HSSFWorkbook workbook, HSSFSheet reportSheet, HSSFSheet basicValuesSheet,
            List<DocumentWrapper> documents, ReportStyle reportStyle, HSSFDataFormat numberFormatter) {
        this.workbook = workbook;
        this.reportSheet = reportSheet;
        this.basicValuesSheet = basicValuesSheet;
        this.documents = documents;
        this.reportStyle = reportStyle;
        this.numberFormatter = numberFormatter;
        this.rowNumber = 0;
        this.columnNumber = 0;
        this.documentCount = documents.size();

        initializeStyles();
    }

    public void create() {
        setupBasicValuesSheet();
        createInternal();
    }

    protected abstract void createInternal();

    protected abstract void initializeStyles();

    //<editor-fold desc="row utilities">

    private Row addRow() {
        return addRow(0);
    }

    private Row addRow(int initialColumnNumber) {
        Row row = reportSheet.createRow(rowNumber);

        rowNumber++;
        columnNumber = initialColumnNumber;

        return row;
    }

    protected void skipRow() {
        rowNumber++;
    }

    //</editor-fold>

    //<editor-fold desc="cell utilities">

    private <T> void addCell(Row row, T value, HSSFCellStyle style) {
        Cell cell = row.createCell(columnNumber);

        if (style != null)
            cell.setCellStyle(style);

        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else {
            cell.setCellValue((String) value);
        }

        columnNumber++;
    }

    private void skipCell() {
        skipCell(1);
    }

    private void skipCell(int skipNumber) {
        columnNumber += skipNumber;
    }

    //</editor-fold>

    //<editor-fold desc="addCompareReportRow">

    protected void addCompareReportRow(String rowName) {
        addCompareReportRow(rowName, null, null, d -> null, null, 0);
    }

    protected <T> void addCompareReportRow(String rowName, Function<DocumentWrapper, T> valueFunction) {
        addCompareReportRow(rowName, valueFunction, null, d -> null, null, 2);
    }

    protected <T> void addCompareReportRow(
            String rowName, Function<DocumentWrapper, T> valueFunction, HSSFCellStyle valueStyle) {
        addCompareReportRow(rowName, valueFunction, CellType.NUMERIC, d -> valueStyle, null, 1);
    }

    protected <T> void addCompareReportRow(
            String rowName, Function<DocumentWrapper, T> valueFunction,
            HSSFCellStyle valueStyle, HSSFCellStyle titleStyle) {
        addCompareReportRow(rowName, valueFunction, CellType.NUMERIC, d -> valueStyle, titleStyle, 1);
    }

    protected <T> void addCompareReportRow(
            String rowName, Function<DocumentWrapper, T> valueFunction,
            CellType cellType, Function<DocumentWrapper, HSSFCellStyle> valueStyleFunction,
            HSSFCellStyle titleStyle, int initialColumnIndex) {
        Row row = addRow(initialColumnIndex);
        Cell titleCell = row.createCell(initialColumnIndex);

        titleCell.setCellValue(rowName);

        if (titleStyle != null)
            titleCell.setCellStyle(titleStyle);

        if (valueFunction != null) {
            int columnIndex = initialColumnIndex + 1;

            for (DocumentWrapper document : documents) {
                Cell cell = cellType == null
                        ? row.createCell(columnIndex)
                        : row.createCell(columnIndex, cellType);

                T value = valueFunction.apply(document);

                if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                } else if (value instanceof Long) {
                    cell.setCellValue((Long) value);
                } else if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else {
                    cell.setCellValue((String) value);
                }

                if (reportStyle != null)
                    cell.setCellStyle(valueStyleFunction.apply(document));

                columnIndex++;
            }
        }
    }

    //</editor-fold>
        }
    }

    protected double getAndereVorderingen(int index) {
        return Double.parseDouble(documents.get(index).getPropertiesMap()

    protected double getAndereVorderingen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.BAVorderingenHoogstens1JaarOverigeVorderingen));
    }

    protected double getAantalWerknemersOpEindeBoekjaar(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBAantalWerknemersOpEindeBoekjaar));
    }

    protected double getAantalBediendenOpEindeBoekjaar(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBAantalBediendenOpEindeBoekjaar));
    }

    protected double getAantalArbeiderssOpEindeBoekjaar(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBAantalArbeidersOpEindeBoekjaar));
    }

    protected double getPersoneelskostenUitzendkrachten(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBPersoneelskostenUitzendkrachten));
    }

    protected double getGepresteerdeUrenUitzendkrachten(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBGepresteerdeUrenUitzendkrachten));
    }

    protected double getGemiddeldeAantalFTEUitzendkrachten(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.SBGemiddeldAantalFTEUitzendkrachten));
    }

    protected double getSBPersoneelsKosten(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.SBPersoneelskosten));
    }

    protected double getGepresteerdeUren(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.SBGepresteerdeUren));
    }

    protected double getGemiddeldeAantalFTE(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.SBGemiddeldeFTE));
    }

    protected double getLiquideMiddelen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BALiquideMiddelen));
    }

    protected double getTotaleActiva(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BATotaalActiva));
    }

    protected double getReserves(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPReserves));
    }

    protected double getOverdragenWinstVerlies(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPOvergedragenWinstVerlies));
    }

    protected double getBedrijfsOpbrengsten(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRBedrijfsopbrengsten));
    }

    protected double getBedrijfsOpbrengstenOmzet(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRBedrijfsopbrengstenOmzet));
    }

    protected double getBedrijfswinstVerlies(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRBedrijfsWinstVerlies));
    }

    protected double getWinstVerliesBoekjaar(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRWinstVerliesBoekjaar));
    }

    protected double getProvisies(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.BPVoorzieningenUitgesteldeBelastingen));
    }

    protected double getFinancieringsLast(DocumentWrapper document) {
        return (getKorteTermijnFinancieleSchulden(document) + getLangeTermijnFinancieleSchulden(document)) / getEBITDA(document);
    }

    protected double getCash(DocumentWrapper document) {
        return getLiquideMiddelen(document)
                + Double.parseDouble(document.getPropertiesMap().get(PropertyName.BAOverlopendeRekeningen));
    }

    protected double getVoorraadrotatie(DocumentWrapper document) {
        return getVoorradenEnBestellingenInUitvoering(document) / getBedrijfsOpbrengsten(document) * 365;
    }

    protected double getKlantenKrediet(DocumentWrapper document) {
        return getHandelsvorderingen(document) / getBedrijfsOpbrengsten(document) * 365;
    }

    protected double getLeveranciersKrediet(DocumentWrapper document) {
        return getLeveranciers(document) / getBedrijfsOpbrengsten(document) * 365;
    }

    protected double getCashFlow(DocumentWrapper document) {
        return getWinstVerliesBoekjaar(document) + getAfschrijvingen(document);
    }

    protected double getWaardeVermindering(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(
                PropertyName.RRBedrijfskostenWaardeverminderingenVoorradenBestellingenUitvoeringHandelsvorderingenToevoegingenTerugnemingen));
    }

    protected double getBelastingen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRBelastingenOpResultaat))
                - Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.RROntrekkingenUitgesteldeBelastingen))
                + Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.RROverboekingUitgesteldeBelastingen));
    }

    protected double getDienstenEnDiverseGoederen(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.RRBedrijfskostenDienstenDiverseGoederen));
    }

    protected double getAndereKosten(DocumentWrapper document) {
        double value = Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.RRBedrijfskostenDienstenDiverseGoederen))
                + Double.parseDouble(document.getPropertiesMap().get(
                PropertyName.RRBedrijfskostenVoorzieningenRisicosKostenToevoegingenBestedingenTerugnemingen))
                + Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.RRBedrijfskostenAndereBedrijfskosten))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten));
        if (Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenUitzonderlijkeKosten)) != Double
                .parseDouble(document.getPropertiesMap()
                        .get(PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten))) {
            value += Double.parseDouble(
                    document.getPropertiesMap().get(PropertyName.RRBedrijfskostenUitzonderlijkeKosten));
        }
        return value;
    }

    protected double getPersoneelskosten(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenBezoldigingenSocialeLastenPensioenen));
    }

    protected double getAankopen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffen));
    }

    protected double getKorteTermijnFinancieleSchulden(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.BPSchuldenHoogstens1JaarFinancieleSchulden));
    }

    protected double getLangeTermijnFinancieleSchulden(DocumentWrapper document) {
        return Double.parseDouble(
                document.getPropertiesMap().get(PropertyName.BPSchuldenMeer1JaarFinancieleSchulden));
    }

    protected double getTotalePassiva(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPTotaalPassiva));
    }

    protected double getKorteTermijnSchulden(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPSchuldenHoogstens1Jaar))
                + Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPOverlopendeRekeningen));
    }

    protected double getVlottendeActiva(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BAVlottendeActiva));
    }

    protected double getInvesteringen(DocumentWrapper document) {
        // NV
        return Double.parseDouble(document.getPropertiesMap().get(
                PropertyName.TLMVAConcessiesOctrooienLicentiesKnowhowMerkenSoortgelijkeRechtenMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLIMVAMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLMVATerreinenEnGebouwenMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLMVAInstallatiesMachinesUitrustingMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLMVAMeubilairRollendMaterieelMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLMVAOverigeMaterieleActivaMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLFVAOndernemingenDeelnemingsverhoudingMutatiesTijdensBoekjaarAanschaffingen))
                // BVBA
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLIMVAMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLMVAMutatiesTijdensBoekjaarAanschaffingen))
                + Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.TLFVAMutatiesTijdensBoekjaarAanschaffingen));
    }

    protected double getAfschrijvingen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(
                PropertyName.RRBedrijfskostenAfschrijvingenWaardeverminderingenOprichtingskostenImmaterieleMaterieleVasteActiva));
    }

    protected double getVasteActiva(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BAVasteActiva));
    }

    protected double getEigenVermogen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPEigenVermogen));
    }

    protected double getLeveranciers(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.BPSchuldenHoogstens1JaarHandelsschuldenLeveranciers));
    }

    protected double getHandelsvorderingen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.BAVorderingenHoogstens1JaarHandelsvorderingen));
    }

    protected double getResultaatVoorBelastingen(DocumentWrapper document) {
        return getEBITDA(document) - getAfschrijvingen(document) - getWaardeVermindering(document)
                + getFinancieleResultaten(document) + getUitzonderlijkeResultaten(document);
    }

    protected double getRRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen));
    }

    protected double getUitzonderlijkeResultaten(DocumentWrapper document) {
        double value = 0;
        if (Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfsopbrengstenUitzonderlijkeOpbrengsten)) != Double
                .parseDouble(document.getPropertiesMap()
                        .get(PropertyName.RRBedrijfsopbrengstenNietRecurrenteBedrijfsopbrengsten))) {
            value += Double.parseDouble(document.getPropertiesMap()
                    .get(PropertyName.RRBedrijfsopbrengstenUitzonderlijkeOpbrengsten));
        }
        if (Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenUitzonderlijkeKosten)) != Double
                .parseDouble(document.getPropertiesMap()
                        .get(PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten))) {
            value -= Double.parseDouble(
                    document.getPropertiesMap().get(PropertyName.RRBedrijfskostenUitzonderlijkeKosten));
        }
        return value;
    }

    protected double getEBITDA(DocumentWrapper document) {
        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA) || reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            return getBrutoMarge(document) - getBedrijfskostenVoorBerekeningen(document);
        } else if (reportStyle.equals(ReportStyle.HISTORIEKNV) || reportStyle.equals(ReportStyle.VERGELIJKINGNV)) {
            return getBedrijfsOpbrengsten(document) - getBedrijfskostenVoorBerekeningen(document);
        }
        return 0;
    }

    protected double getEBIT(DocumentWrapper document) {
        return getEBITDA(document) - getAfschrijvingen(document) - getWaardeVermindering(document);
    }

    protected double getFinancieleResultaten(DocumentWrapper document) {
        double inkom = Double
                .parseDouble(document.getPropertiesMap().get(PropertyName.RRFinancieleOpbrengsten));
        if (inkom == 0) {
            inkom = Double.parseDouble(
                    document.getPropertiesMap().get(PropertyName.RRFinancieleOpbrengstenRecurrent));
        }
        double uitgaand = getFinancieleKosten(document);
        if (uitgaand == 0) {
            uitgaand = Double
                    .parseDouble(document.getPropertiesMap().get(PropertyName.RRFinancieleKostenRecurrent));
        }
        return inkom - uitgaand;
    }

    protected double getFinancieleKosten(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.RRFinancieleKosten));
    }

    protected double getCoreNettoWerkKapitaal(DocumentWrapper document) {
        return getVoorradenEnBestellingenInUitvoering(document) + getHandelsvorderingen(document) - getLeveranciers(document);
    }

    protected double getCapitalEmployed(DocumentWrapper document) {
        return getCoreNettoWerkKapitaal(document) + getVasteActiva(document);
    }

    protected double getBedrijfskostenVoorBerekeningen(DocumentWrapper document) {
        return getAankopen(document) + getPersoneelskosten(document) + getAndereKosten(document);
    }

    protected double getTotaleBedrijfskosten(DocumentWrapper document) {
        return Double.parseDouble((document.getPropertiesMap().get(PropertyName.RRBedrijfskosten)));
    }

    protected double getBrutoMarge(DocumentWrapper document) {
        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA) || reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BVBABrutomarge));
        } else {
            return getBedrijfsOpbrengsten(document) - getNietRecurenteBedrijfsopbrengsten(document)
                    - getRRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen(document)
                    - getBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename(document)
                    - getDienstenEnDiverseGoederen(document);
        }
    }

    protected double getBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename));
    }

    protected double getNietRecurenteBedrijfsopbrengsten(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.RRBedrijfsopbrengstenNietRecurrenteBedrijfsopbrengsten));
    }

    protected double getToegevoegdeWaarde(DocumentWrapper document) {
        return getBedrijfsOpbrengstenOmzet(document) - getAankopen(document);
    }

    protected double getAndereSchuldenKorteTermijn(DocumentWrapper document) throws NumberFormatException {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPSchuldenHoogstens1Jaar))
                + Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPOverlopendeRekeningen))
                - getLeveranciers(document) - Double.parseDouble(document.getPropertiesMap()
                .get(PropertyName.BPSchuldenHoogstens1JaarFinancieleSchulden));
    }

    protected double getVoorradenEnBestellingenInUitvoering(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BAVoorradenBestellingenUitvoering));
    }

    protected double getTotaleSchulden(DocumentWrapper document) {
        return Double.parseDouble(document.getPropertiesMap().get(PropertyName.BPSchulden));
    }

    protected double getNettoWerkkapitaal(DocumentWrapper document) {
        return getVoorradenEnBestellingenInUitvoering(document) + getHandelsvorderingen(document) - getLeveranciers(document);
    }

    protected double getZScoreAltman(DocumentWrapper document) {
        double x1 = getNettoWerkkapitaal(document) / getTotaleActiva(document) * 0.717;
        double x2 = getWinstVerliesBoekjaar(document) / getTotaleActiva(document) * 0.847;
        double x3 = getEBIT(document) / getTotaleActiva(document) * 3.107;
        double x4 = getEigenVermogen(document) / getTotaleSchulden(document) * 0.42;
        double x5 = getBedrijfsOpbrengstenOmzet(document) / getTotaleActiva(document) * 0.998;

        return x1 + x2 + x3 + x4 + x5;
    }
}
