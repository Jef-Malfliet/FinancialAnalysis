package Models.Interfaces;

import Models.DocumentWrapper;

public interface IDocumentBuilder {

    String getName();

    boolean isSelected();

    DocumentWrapper build();

    //BALANS NA WINSTVERDELING
    //ACTIVA

    //VASTE ACTIVA
    IDocumentBuilder addBAVasteActiva();

    //IMMATERIËLE ACTIVA
    IDocumentBuilder addBAImmaterieleVasteActiva();

    //MATERIËLE ACTIVA
    IDocumentBuilder addBAMaterieleVasteActiva();

    //FINANCIËLE ACTIVA
    IDocumentBuilder addBAFinancieleVasteActiva();

    //VLOTTENDE ACTIVA
    IDocumentBuilder addBAVlottendeActiva();

    IDocumentBuilder addBAVoorradenBestellingenUitvoering();

    IDocumentBuilder addBAVorderingenHoogstens1JaarHandelsvorderingen();

    IDocumentBuilder addBAVorderingenHoogstens1JaarOverigeVorderingen();

    IDocumentBuilder addBALiquideMiddelen();

    IDocumentBuilder addBAOverlopendeRekeningen();

    IDocumentBuilder addBATotaalActiva();

    //PASSIVA
    //EIGEN VERMOGEN
    IDocumentBuilder addBPEigenVermogen();

    IDocumentBuilder addBPReserves();

    IDocumentBuilder addBPOvergedragenWinstVerlies();

    //VOORZIENINGEN EN UITGESTELDE BELASTINGEN
    IDocumentBuilder addBPVoorzieningenUitgesteldeBelastingen();

    IDocumentBuilder addBPSchulden();

    IDocumentBuilder addBPSchuldenMeer1Jaar();

    IDocumentBuilder addBPSchuldenMeer1JaarFinancieleSchulden();

    IDocumentBuilder addBPSchuldenMeer1JaarOverigeSchulden();

    IDocumentBuilder addBPSchuldenHoogstens1Jaar();

    IDocumentBuilder addBPSchuldenHoogstens1JaarFinancieleSchulden();

    IDocumentBuilder addBPSchuldenHoogstens1JaarHandelsschuldenLeveranciers();

    IDocumentBuilder addBPOverlopendeRekeningen();

    IDocumentBuilder addBPTotaalPassiva();

    //RESULATATENREKENING
    //BEDRIJFOPBRENGST
    IDocumentBuilder addRRBedrijfsopbrengsten();

    IDocumentBuilder addRRBedrijfsopbrengstenOmzet();

    IDocumentBuilder addRRBedrijfsopbrengstenNietRecurrenteBedrijfsopbrengsten();

    //BEDRIJFSKOSTEN
    IDocumentBuilder addRRBedrijfskosten();

    IDocumentBuilder addRRBedrijfskostenHandelsgoederenGrondHulpstoffen();

    IDocumentBuilder addRRBedrijfskostenHandelsgoederenGrondHulpstoffenAankopen();

    IDocumentBuilder addRRBedrijfskostenHandelsgoederenGrondHulpstoffenVoorraadAfnameToename();

    IDocumentBuilder addRRBedrijfskostenDienstenDiverseGoederen();

    IDocumentBuilder addRRBedrijfskostenBezoldigingenSocialeLastenPensioenen();

    IDocumentBuilder addRRBedrijfskostenAfschrijvingenWaardeverminderingenOprichtingskostenImmaterieleMaterieleVasteActiva();

    IDocumentBuilder addRRBedrijfskostenWaardeverminderingenVoorradenBestellingenUitvoeringHandelsvorderingenToevoegingenTerugnemingen();

    IDocumentBuilder addRRBedrijfskostenVoorzieningenRisicosKostenToevoegingenBestedingenTerugnemingen();

    IDocumentBuilder addRRBedrijfskostenAndereBedrijfskosten();

    IDocumentBuilder addRRBedrijfskostenNietRecurrenteBedrijfskosten();

    IDocumentBuilder addRRBedrijfskostenUitzonderlijkeKosten();

    IDocumentBuilder addRRBedrijfsopbrengstenUitzonderlijkeOpbrengsten();

    //BEDRIJFSWINSTVERLIES
    IDocumentBuilder addRRBedrijfsWinstVerlies();

    //FINANCIËLE OPBRENGSTEN
    IDocumentBuilder addRRFinancieleOpbrengsten();

    IDocumentBuilder addRRFinancieleOpbrengstenRecurrent();

    //FINANCIËLE KOSTEN
    IDocumentBuilder addRRFinancieleKosten();

    IDocumentBuilder addRRFinancieleKostenRecurrent();

    //ANDERE

    IDocumentBuilder addRROntrekkingenUitgesteldeBelastingen();

    IDocumentBuilder addRROverboekingUitgesteldeBelastingen();

    IDocumentBuilder addRRBelastingenOpResultaat();

    IDocumentBuilder addRRWinstVerliesBoekjaar();

    IDocumentBuilder addTLMVAMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLIMVAMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLMVAConcessiesOctrooienLicentiesKnowhowMerkenSoortgelijkeRechtenMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLFVAMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLMVATerreinenEnGebouwenMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLMVAInstallatiesMachinesUitrustingMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLMVAMeubilairRollendMaterieelMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLMVAOverigeMaterieleActivaMutatiesTijdensBoekjaarAanschaffingen();

    IDocumentBuilder addTLFVAOndernemingenDeelnemingsverhoudingMutatiesTijdensBoekjaarAanschaffingen();

    //SOCIALE BALANS
    IDocumentBuilder addSBGemiddeldeFTE();

    IDocumentBuilder addSBGepresteerdeUren();

    IDocumentBuilder addSBGemiddeldAantalFTEUitzendkrachten();

    IDocumentBuilder addSBGepresteerdeUrenUitzendkrachten();

    IDocumentBuilder addSBPersoneelskostenUitzendkrachten();

    IDocumentBuilder addSBAantalWerknemersOpEindeBoekjaar();

    IDocumentBuilder addSBAantalBediendenOpEindeBoekjaar();

    IDocumentBuilder addSBAantalArbeidersOpEindeBoekjaar();

    IDocumentBuilder addBVBABrutomarge();
}
