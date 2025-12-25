package Services;

import Models.Enums.PropertyName;
import Models.Enums.ReportStyle;
import Models.DocumentWrapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;
import java.util.function.Function;

public abstract class ReportService {

    protected final XSSFSheet reportSheet;
    private final XSSFSheet basicValuesSheet;
    private final XSSFFormulaEvaluator formulaEvaluator;
    private final XSSFSheetConditionalFormatting conditionalFormatting;
    private XSSFTable basicValuesTable;

    protected final List<DocumentWrapper> documents;
    protected final int documentCount;
    protected final ReportStyle reportStyle;
    protected final XSSFWorkbook workbook;
    protected final XSSFDataFormat numberFormatter;
    protected int rowNumber;
    protected int columnNumber;

    public ReportService(XSSFWorkbook workbook, XSSFSheet reportSheet, XSSFSheet basicValuesSheet, List<DocumentWrapper> documents, ReportStyle reportStyle, XSSFDataFormat numberFormatter) {
        this.workbook = workbook;
        this.reportSheet = reportSheet;
        this.formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        this.conditionalFormatting = reportSheet.getSheetConditionalFormatting();

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
        resetRowNumber();
        setupBasicValuesSheet();
        resetRowNumber();
        createInternal();
    }

    protected abstract void createInternal();

    protected abstract void initializeStyles();

    //<editor-fold desc="row utilities">

    protected XSSFRow addRow(XSSFSheet sheet) {
        return addRow(sheet, 0);
    }

    protected XSSFRow addRow(XSSFSheet sheet, int initialColumnNumber) {
        XSSFRow row = sheet.createRow(rowNumber);

        rowNumber++;
        columnNumber = initialColumnNumber;

        return row;
    }

    protected void skipRow() {
        rowNumber++;
    }

    protected void resetRowNumber() {
        rowNumber = 0;
    }

    //</editor-fold>

    //<editor-fold desc="cell utilities">

    protected <T> void addCell(XSSFRow row, T value) {
        addCell(row, value, null);
    }

    protected <T> void addCell(XSSFRow row, T value, XSSFCellStyle style) {
        Cell cell = row.createCell(columnNumber);

        if (style != null) cell.setCellStyle(style);

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

    protected void addReferenceCell(XSSFRow row, int cellIndex, String formula, CellType cellType) {
        addReferenceCell(row, cellIndex, formula, null, cellType);
    }

    protected void addReferenceCell(XSSFRow row, String formula, XSSFCellStyle style) {
        addReferenceCell(row, columnNumber, formula, style, null);
    }

    protected void addReferenceCell(XSSFRow row, int index, String formula, XSSFCellStyle style, CellType type) {
        Cell cell = type == null ? row.createCell(index) : row.createCell(index, type);

        if (style != null) cell.setCellStyle(style);

        if (formula == null || formula.equals("/")) {
            cell.setCellValue(formula);
        } else {
            cell.setCellFormula(formula);
        }

        formulaEvaluator.evaluateFormulaCell(cell);

        columnNumber++;
    }

    protected void skipCell() {
        skipCell(1);
    }

    protected void skipCell(int skipNumber) {
        columnNumber += skipNumber;
    }

    //</editor-fold>

    private void setupBasicValuesSheet() {
        AreaReference areaReference = workbook.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(documentCount + 1, PropertyName.values().length));

        basicValuesTable = basicValuesSheet.createTable(areaReference);
        basicValuesTable.setName("Values");
        basicValuesTable.setDisplayName("Values");
        basicValuesTable.getCTTable().addNewTableStyleInfo();
        basicValuesTable.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");

        setupBasicValueTableHeader();
        setupBasicValueTableValues();
    }

    private void setupBasicValueTableValues() {
        for (DocumentWrapper document : documents) {
            XSSFRow documentRow = addRow(basicValuesSheet);

            for (PropertyName propertyName : PropertyName.values()) {
                double value = Double.parseDouble(document.getPropertiesMap().get(propertyName));

                if (propertyName.mustBeRounded()) {
                    addCell(documentRow, Math.round(value));
                } else {
                    addCell(documentRow, value);
                }
            }
        }
    }

    private void setupBasicValueTableHeader() {
        XSSFRow headerRow = addRow(basicValuesSheet);

        for (PropertyName propertyName : PropertyName.values()) {
            addCell(headerRow, propertyName.toString());
        }

        basicValuesTable.updateHeaders();
    }

    // <editor-fold desc="getters">

    protected String getBasicValueReference(int index, PropertyName propertyName) {
        return String.format("INDEX(%s[%s], %d)", basicValuesTable.getName(), propertyName.toString(), index + 1);
        // return "'" + basicValuesSheet.getSheetName() + "'!" + basicValuesTable.getName() + "[" + intToExcelColumn(index) + (propertyName.ordinal() + 1);
    }

    protected double getBasicValue(int index, PropertyName propertyName) {
        return basicValuesSheet.getRow(index + 1).getCell(propertyName.ordinal()).getNumericCellValue();
    }

    protected String getYear(int i) {
        return String.valueOf(documents.get(i).getYear());
    }

    protected String getVorderingenHoogstens1JaarOverigeVorderingen(int i) {
        return getBasicValueReference(i, PropertyName.BAVorderingenHoogstens1JaarOverigeVorderingen);
    }

    protected String getGemiddeldeAantalFTEUitzendkrachten(int i) {
        return getBasicValueReference(i, PropertyName.SBGemiddeldAantalFTEUitzendkrachten);
    }

    protected String getGepresteerdeUrenUitzendkrachten(int i) {
        return getBasicValueReference(i, PropertyName.SBGepresteerdeUrenUitzendkrachten);
    }

    protected String getPersoneelskostenUitzendkrachten(int i) {
        return getBasicValueReference(i, PropertyName.SBPersoneelskostenUitzendkrachten);
    }

    protected String getAantalWerknemersOpEindeBoekjaar(int i) {
        return getBasicValueReference(i, PropertyName.SBAantalWerknemersOpEindeBoekjaar);
    }

    protected String getAantalBediendenOpEindeBoekjaar(int i) {
        return getBasicValueReference(i, PropertyName.SBAantalBediendenOpEindeBoekjaar);
    }

    protected String getAantalArbeidersOpEindeBoekjaar(int i) {
        return getBasicValueReference(i, PropertyName.SBAantalArbeidersOpEindeBoekjaar);
    }

    protected String getGepresteerdeUren(int i) {
        return getBasicValueReference(i, PropertyName.SBGepresteerdeUren);
    }

    protected String getTotaleActiva(int i) {
        return getBasicValueReference(i, PropertyName.BATotaalActiva);
    }

    protected String getVoorradenBestellingenUitvoering(int i) {
        return getBasicValueReference(i, PropertyName.BAVoorradenBestellingenUitvoering);
    }

    protected String getGemiddeldeAantalFTE(int i) {
        return getBasicValueReference(i, PropertyName.SBGemiddeldeFTE);
    }

    protected String getBedrijfsOpbrengsten(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfsopbrengsten);
    }

    protected String getLiquideMiddelen(int i) {
        return getBasicValueReference(i, PropertyName.BALiquideMiddelen);
    }

    protected String getBedrijfsOpbrengstenOmzet(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfsopbrengstenOmzet);
    }

    protected String getWinstVerliesBoekjaar(int i) {
        return getBasicValueReference(i, PropertyName.RRWinstVerliesBoekjaar);
    }

    protected String getReserves(int i) {
        return getBasicValueReference(i, PropertyName.BPReserves);
    }

    protected String getOverdragenWinstVerlies(int i) {
        return getBasicValueReference(i, PropertyName.BPOvergedragenWinstVerlies);
    }

    protected String getBedrijfsWinstVerlies(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfsWinstVerlies);
    }

    protected String getProvisies(int i) {
        return getBasicValueReference(i, PropertyName.BPVoorzieningenUitgesteldeBelastingen);
    }

    protected String getLangeTermijnSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenMeer1Jaar);
    }

    protected String getFinancieringsLast(int i) {
        return getKorteTermijnFinancieleSchulden(i)
                .add(getLangeTermijnFinancieleSchulden(i))
                .inParenthesis()
                .dividedBy(getEBITDA(i)
                        .inParenthesis());
    }

    protected String getCash(int i) {
        return getLiquideMiddelen(i)
                .add(getBasicValueReference(i, PropertyName.BAOverlopendeRekeningen));
    }

    protected String getVoorraadrotatie(int i) {
        return getVoorradenBestellingenUitvoering(i)
                .dividedBy(getBedrijfsOpbrengsten(i))
                .multipliedBy("365");
    }

    protected String getKlantenKrediet(int i) {
        return getHandelsvorderingen(i)
                .dividedBy(getBedrijfsOpbrengsten(i))
                .multipliedBy("365");
    }

    protected String getLeveranciersKrediet(int i) {
        return getLeveranciers(i)
                .dividedBy(getBedrijfsOpbrengsten(i))
                .multipliedBy("365");
    }

    protected String getCashFlow(int i) {
        return getWinstVerliesBoekjaar(i)
                .add(getAfschrijvingen(i));
    }

    protected String getLiquiditeitsRatio(int i) {
        return getVlottendeActiva(i)
                .dividedBy(getKorteTermijnSchulden(i)
                        .inParenthesis());
    }

    protected String getSolvabiliteitsRatio(int i) {
        return getKorteTermijnSchulden(i)
                .inParenthesis()
                .dividedBy(getTotaleActiva(i));
    }

    protected String getNettoWinstOverOmzet(int i) {
        return getWinstVerliesBoekjaar(i)
                .dividedBy(getBedrijfsOpbrengstenOmzet(i));
    }

    protected String getRentabiliteitsRatioEigenVermogen(int i) {
        return getWinstVerliesBoekjaar(i)
                .dividedBy(getEigenVermogen(i));
    }

    protected String getCashConversionCycle(int i) {
        return getVoorraadrotatie(i)
                .add(getKlantenKrediet(i))
                .subtract(getLeveranciersKrediet(i));
    }

    protected String getWaardeVermindering(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenWaardeverminderingenVoorradenBestellingenUitvoeringHandelsvorderingenToevoegingenTerugnemingen);
    }

    protected String getBelastingen(int i) {
        return getBasicValueReference(i, PropertyName.RRBelastingenOpResultaat)
                .subtract(getBasicValueReference(i, PropertyName.RROntrekkingenUitgesteldeBelastingen))
                .add(getBasicValueReference(i, PropertyName.RROverboekingUitgesteldeBelastingen));
    }

    protected String getDienstenEnDiverseGoederen(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenDienstenDiverseGoederen);
    }

    protected String getAndereKosten(int i) {
        String value = getDienstenEnDiverseGoederen(i)
                .add(getBasicValueReference(i, PropertyName.RRBedrijfskostenVoorzieningenRisicosKostenToevoegingenBestedingenTerugnemingen))
                .add(getBasicValueReference(i, PropertyName.RRBedrijfskostenAndereBedrijfskosten))
                .add(getBasicValueReference(i, PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten));

        if (getBasicValue(i, PropertyName.RRBedrijfskostenUitzonderlijkeKosten) == getBasicValue(i, PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten)) {
            return value;
        }

        return value.add(getBasicValueReference(i, PropertyName.RRBedrijfskostenUitzonderlijkeKosten));
    }

    protected String getPersoneelskosten(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenBezoldigingenSocialeLastenPensioenen);
    }

    protected String getAankopen(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffen);
    }

    protected String getKorteTermijnFinancieleSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenHoogstens1JaarFinancieleSchulden);
    }

    protected String getLangeTermijnFinancieleSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenMeer1JaarFinancieleSchulden);
    }

    protected String getLangeTermijnOverigeSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenMeer1JaarOverigeSchulden);
    }

    protected String getTotalePassiva(int i) {
        return getBasicValueReference(i, PropertyName.BPTotaalPassiva);
    }

    protected String getKorteTermijnSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenHoogstens1Jaar)
                .add(getBasicValueReference(i, PropertyName.BPOverlopendeRekeningen));
    }

    protected String getVlottendeActiva(int i) {
        return getBasicValueReference(i, PropertyName.BAVlottendeActiva);
    }

    protected String getInvesteringen(int i) {
        // NV
        return getBasicValueReference(i, PropertyName.TLMVAConcessiesOctrooienLicentiesKnowhowMerkenSoortgelijkeRechtenMutatiesTijdensBoekjaarAanschaffingen)
                .add(getBasicValueReference(i, PropertyName.TLIMVAMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLMVATerreinenEnGebouwenMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLMVAInstallatiesMachinesUitrustingMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLMVAMeubilairRollendMaterieelMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLMVAOverigeMaterieleActivaMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLFVAOndernemingenDeelnemingsverhoudingMutatiesTijdensBoekjaarAanschaffingen))
                // BVBA
                .add(getBasicValueReference(i, PropertyName.TLIMVAMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLMVAMutatiesTijdensBoekjaarAanschaffingen))
                .add(getBasicValueReference(i, PropertyName.TLFVAMutatiesTijdensBoekjaarAanschaffingen));
    }

    protected String getAfschrijvingen(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenAfschrijvingenWaardeverminderingenOprichtingskostenImmaterieleMaterieleVasteActiva);
    }

    protected String getVasteActiva(int i) {
        return getBasicValueReference(i, PropertyName.BAVasteActiva);
    }

    protected String getImmaterieleVasteActiva(int i) {
        return getBasicValueReference(i, PropertyName.BAImmaterieleVasteActiva);
    }

    protected String getMaterieleVasteActiva(int i) {
        return getBasicValueReference(i, PropertyName.BAMaterieleVasteActiva);
    }

    protected String getFinancieleVasteActiva(int i) {
        return getBasicValueReference(i, PropertyName.BAFinancieleVasteActiva);
    }

    protected String getEigenVermogen(int i) {
        return getBasicValueReference(i, PropertyName.BPEigenVermogen);
    }

    protected String getLeveranciers(int i) {
        return getBasicValueReference(i, PropertyName.BPSchuldenHoogstens1JaarHandelsschuldenLeveranciers);
    }

    protected String getHandelsvorderingen(int i) {
        return getBasicValueReference(i, PropertyName.BAVorderingenHoogstens1JaarHandelsvorderingen);
    }

    protected String getResultaatVoorBelastingen(int i) {
        return getEBITDA(i)
                .subtract(getAfschrijvingen(i))
                .subtract(getWaardeVermindering(i))
                .add(getFinancieleResultaten(i))
                .add(getUitzonderlijkeResultaten(i));
    }

    protected String getRRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen);
    }

    protected String getUitzonderlijkeResultaten(int i) {
        String value = "";

        if (getBasicValue(i, PropertyName.RRBedrijfsopbrengstenUitzonderlijkeOpbrengsten) != getBasicValue(i, PropertyName.RRBedrijfsopbrengstenNietRecurrenteBedrijfsopbrengsten)) {
            value = getBasicValueReference(i, PropertyName.RRBedrijfsopbrengstenUitzonderlijkeOpbrengsten);
        }

        if (getBasicValue(i, PropertyName.RRBedrijfskostenUitzonderlijkeKosten) != getBasicValue(i, PropertyName.RRBedrijfskostenNietRecurrenteBedrijfskosten)) {
            value = value.subtract(getBasicValueReference(i, PropertyName.RRBedrijfskostenUitzonderlijkeKosten));
        }

        if (value.isEmpty())
            return "0";

        return value;
    }

    protected String getEBITDA(int i) {
        switch (reportStyle) {
            case HISTORIEKBVBA:
            case VERGELIJKINGBVBA:
                return getBrutoMarge(i)
                        .subtract(getBedrijfskostenVoorBerekeningen(i)
                                .inParenthesis());
            case HISTORIEKNV:
            case VERGELIJKINGNV:
                return getBedrijfsOpbrengsten(i)
                        .subtract(getBedrijfskostenVoorBerekeningen(i)
                                .inParenthesis());
            default:
                return "0";
        }
    }

    protected String getEBIT(int i) {
        return getEBITDA(i)
                .subtract(getAfschrijvingen(i))
                .subtract(getWaardeVermindering(i));
    }

    protected String getFinancieleResultaten(int i) {
        String profit = getBasicValueReference(i, PropertyName.RRFinancieleOpbrengsten);

        if (profit.equals("0")) {
            profit = getBasicValueReference(i, PropertyName.RRFinancieleOpbrengstenRecurrent);
        }

        String costs = getFinancieleKosten(i);

        if (costs.equals("0")) {
            costs = getBasicValueReference(i, PropertyName.RRFinancieleKostenRecurrent);
        }

        return profit
                .subtract(costs);
    }

    protected String getFinancieleKosten(int i) {
        return getBasicValueReference(i, PropertyName.RRFinancieleKosten);
    }

    protected String getCoreNettoWerkKapitaal(int i) {
        return getVoorradenBestellingenUitvoering(i)
                .add(getHandelsvorderingen(i))
                .subtract(getLeveranciers(i));
    }

    protected String getCapitalEmployed(int i) {
        return getCoreNettoWerkKapitaal(i)
                .add(getVasteActiva(i));
    }

    protected String getBedrijfskostenVoorBerekeningen(int i) {
        return getAankopen(i)
                .add(getPersoneelskosten(i))
                .add(getAndereKosten(i));
    }

    protected String getTotaleBedrijfskosten(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskosten);
    }

    protected String getBrutoMarge(int i) {
        if (reportStyle.equals(ReportStyle.HISTORIEKBVBA) || reportStyle.equals(ReportStyle.VERGELIJKINGBVBA)) {
            return getBasicValueReference(i, PropertyName.BVBABrutomarge);
        } else {
            return getBedrijfsOpbrengsten(i)
                    .subtract(getNietRecurenteBedrijfsopbrengsten(i))
                    .subtract(getRRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen(i))
                    .subtract(getBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename(i))
                    .subtract(getDienstenEnDiverseGoederen(i));
        }
    }

    protected String getBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename);
    }

    protected String getNietRecurenteBedrijfsopbrengsten(int i) {
        return getBasicValueReference(i, PropertyName.RRBedrijfsopbrengstenNietRecurrenteBedrijfsopbrengsten);
    }

    protected String getToegevoegdeWaarde(int i) {
        return getBedrijfsOpbrengstenOmzet(i)
                .subtract(getAankopen(i));
    }

    protected String getAndereSchuldenKorteTermijn(int i) throws NumberFormatException {
        return getBasicValueReference(i, PropertyName.BPSchuldenHoogstens1Jaar)
                .add(getBasicValueReference(i, PropertyName.BPOverlopendeRekeningen))
                .subtract(getLeveranciers(i))
                .subtract(getBasicValueReference(i, PropertyName.BPSchuldenHoogstens1JaarFinancieleSchulden));
    }

    protected String getTotaleSchulden(int i) {
        return getBasicValueReference(i, PropertyName.BPSchulden);
    }

    protected String getNettoWerkkapitaal(int i) {
        return getVoorradenBestellingenUitvoering(i)
                .add(getHandelsvorderingen(i))
                .subtract(getLeveranciers(i));
    }

    protected String getZScoreAltman(int i) {
        String x1 = getNettoWerkkapitaal(i)
                .inParenthesis()
                .dividedBy(getTotaleActiva(i)
                        .multipliedBy("0.717")
                        .inParenthesis());
        String x2 = getWinstVerliesBoekjaar(i)
                .dividedBy(getTotaleActiva(i)
                        .multipliedBy("0.847")
                        .inParenthesis());
        String x3 = getEBIT(i)
                .inParenthesis()
                .dividedBy(getTotaleActiva(i)
                        .multipliedBy("3.107")
                        .inParenthesis());
        String x4 = getEigenVermogen(i)
                .dividedBy(getTotaleSchulden(i)
                        .multipliedBy("0.42")
                        .inParenthesis());
        String x5 = getBedrijfsOpbrengstenOmzet(i)
                .dividedBy(getTotaleActiva(i)
                        .multipliedBy("0.998")
                        .inParenthesis());

        return x1.add(x2).add(x3).add(x4).add(x5);
    }

    // </editor-fold>

    protected ConditionalFormattingRule createConditionalFormattingRule(byte operator, String value) {
        return createConditionalFormattingRule(operator, value, null);
    }

    protected ConditionalFormattingRule createConditionalFormattingRule(byte operator, String value1, String value2) {
        return conditionalFormatting.createConditionalFormattingRule(operator, value1, value2);
    }

    protected void addConditionalFormatting(CellRangeAddress[] regions, ConditionalFormattingRule[] cfRules) {
        conditionalFormatting.addConditionalFormatting(regions, cfRules);
    }

    protected String getCurrentCellAsRange() {
        StringBuilder columnName = new StringBuilder();
        int columnNumberCopy = columnNumber;

        if (columnNumberCopy == 0) columnName = new StringBuilder("A");
        else {
            while (columnNumberCopy > 0) {
                int modulo = columnNumberCopy % 26;
                columnName.append('A' + modulo);
                columnNumberCopy = (columnNumberCopy - modulo) / 26;
            }
        }

        return String.format("%s%d:%s%d", columnName, rowNumber, columnName, rowNumber);
    }
}
