using com.salescloud as db from '../db/schema';

// ── Input types ─────────────────────────────────────────────────────────────

type InvolvedPartiesinput {
  buspartner : String;
  role       : String;
}

type ProviderContractinput {
  product        : String;
  portfolio      : String;
  fixedMWh       : String;
  fixedPrice     : String;
  portfolioprice : String;
  subspercent    : String;
  startDate      : String;
  endDate        : String;
  estUsage       : String;
  recPrice       : String;
  migpPercent    : String;
  recpercent     : String;
  migpMWh        : String;
  recMWh         : String;
  enrollChk      : String;   
  exportDate     : String; 
  exportChk : String;
  netPremium     : String; // SMB Fixed $ — Net Premium $/MWh
}

type ConsumptionDetailinput {
  contacc             : String;
  buspartner          : String;
  validity            : String;
  replacementCA       : String;
  est12month          : String;
  est12monthMetered   : String; // ← ADD
  est12monthUnMetered : String; // ← ADD
  RECBlender_flag     : String; // ← ADD
  providercontractId  : String;
  cycle20             : String;
  metered             : String;
  arrear60            : String;
  nsfFlag             : String;
  inactive            : String;
  enrolled            : String;
  selected            : String;
}

type OpportunityInput {
  Oppid                    : String;
  quoteId                  : String;
  objectStatus             : String;
  oppCategory              : String;
  status                   : String;
  annualGross              : String;
  sizeofOpp                : String;
  percentUsage             : String;
  annualSubs               : String;
  oppSubs                  : String;
  contLifetimeval          : String;
  salesType                : String;
  priceStructure           : String;
  nte                      : String;
  annualMWh                : String;
  term                     : String;
  meteredCon               : String;
  unmeteredCon             : String;
  commencementLetterSent   : String;
  commencementLetterSigned : String;
  enrolled                 : String;
  includeInPipeline        : String;
  salesPhase : String;
  enrollmentReferralCode   : String;   // SMB only
  enrollmentEmailId        : String;   // SMB only
}

// ── CloudService ─────────────────────────────────────────────────────────────

service CloudService {

  // Entities for Fiori UI
  entity Opportunity       as
    projection on db.Opportunity
    excluding {
      InvolvedParties,
      ConsumptionDetail,
      ProviderContract,
      Prospect
    };

  entity InvolvedParties   as projection on db.InvolvedParties;
  entity ProviderContract  as projection on db.ProviderContract;
  entity ConsumptionDetail as projection on db.ConsumptionDetail;
  entity Prospect          as projection on db.Prospect;

  @cds.redirection.target: false
  entity OpportunityIds    as select distinct key Oppid from db.Opportunity;

  // Functions & Actions
  function userDetails()                                                       returns String;
  function getOpportunityRecord(connection_object: String)                     returns String;
  action   saveOpportunity(opp: OpportunityInput)                              returns String;
  action   saveFullOpportunity(bundle: String)                                 returns String;
  action   getLatestCPIRecord()                                                returns String;
  action   getAllOppIds()                                                      returns String;
  action fetchValidateCAs(                // ✅ already added
    // businessPartnerId : String,
    // contractAccounts  : String,
    businessPartners : String,    // ← matches oOp.setParameter("businessPartners", ...)
    selectedCAs      : String,    // ← matches oOp.setParameter("selectedCAs", ...)
    oppId : String,
    objectStatus      : String,   // ← ADD
    quoteId           : String    // ← ADD
  )                                       returns String;
  action getOrReplicateQuote(             // ✅ ADD NEW
    oppId   : String,
    quoteId : String
  )                                        returns String;
  
  action getQuoteByQuoteId(quoteId: String)     returns String;  // ← ADD                                    
  action pushSalesPhaseToC4C(bundle: String)    returns String;  // ✅ ADD
  action initiateEnrollment(bundle: String)     returns String;  // ✅ ADD
  action changeBillingCycle(bundle: String)     returns String;  // ✅ ADD
  action updateLastFetchedAt(oppId: String, objectStatus: String, quoteId: String)     returns String;
  action saveExportDate(bundle: String) returns String;
}

// ── CPIService ────────────────────────────────────────────────────────────────

service CPIService @(path: 'cpi') {

  action PushOpportunityFromCPI(opp: OpportunityInput,
                                involvedParties: many InvolvedPartiesinput,
                                providerContracts: many ProviderContractinput) returns {
    status          : String;
    inserted        : Integer;
    updated         : Integer;
    childrenWritten : Integer;
  };
}
