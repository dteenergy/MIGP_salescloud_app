namespace com.salescloud;

using {
  cuid,
  managed
} from '@sap/cds/common';

entity InvolvedParties : cuid, managed {
  buspartner  : String;
  role        : String;
  opportunity : Association to Opportunity;
}

entity ConsumptionDetail : cuid, managed {
  contacc             : String;
  buspartner          : String;
  validity            : String;
  replacementCA       : String;
  est12month          : String;
  providercontractId  : String;
  cycle20             : String;
  metered             : String;
  arrear60            : String;
  nsfFlag             : String;
  inactive            : String;
  enrolled            : String; // FDD #13 — free-text from CR&B, comma-sep e.g. "MIGP-FLEX,LCVP"
  selected            : String; // "true"/"false" — user selection flag
  est12monthMetered   : String; //added on 18-03-26
  est12monthUnMetered : String; //added on 18-03-26
  RECBlender_flag     : String; //added on 18-03-26
  opportunity         : Association to Opportunity;
}

entity ProviderContract : cuid, managed {
  product        : String;
  productId       : String;  
  portfolio      : String;
  fixedMWh       : String;
  fixedPrice     : String; // SMB fixed dollar amount
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
  enrollChk      : String; // BTP suggestion #4 — Enroll checkbox per row ("true"/"false")
  exportDate     : String; // BTP suggestion #4 — auto-stamped when Export File is clicked
  opportunity    : Association to Opportunity;
  exportChk      : String;
  netPremium     : String; // SMB Fixed $ — Net Premium $/MWh (used for Annual Subscribed Load calc)
}

entity Prospect : cuid, managed {
  siteAddLoc   : String;
  projectedCon : String;
  year         : String;
  opportunity  : Association to Opportunity;
}

entity Opportunity : cuid, managed {
  Oppid                    : String; //removed key
  InvolvedParties          : Composition of many InvolvedParties
                               on InvolvedParties.opportunity = $self;
  ConsumptionDetail        : Composition of many ConsumptionDetail
                               on ConsumptionDetail.opportunity = $self;
  ProviderContract         : Composition of many ProviderContract
                               on ProviderContract.opportunity = $self;
  Prospect                 : Composition of many Prospect
                               on Prospect.opportunity = $self;
  quoteId                  : String; //earlier contractid
  oppCategory              : String;
  status                   : String;
  annualGross              : String; // FDD #1  Gross Annual Usage
  percentUsage             : String; // FDD #2  Estimated Subscription %
  sizeofOpp                : String; // FDD #3  Size of Opportunity
  annualMWh                : String; // FDD #4  Annual MWh at Close
  annualSubs               : String; // FDD #5  Annual Subscribed Load
  priceStructure           : String; // FDD #6
  salesType                : String; // FDD #7
  contLifetimeval          : String; // FDD #8  Contract Lifetime Value
  oppSubs                  : String; // FDD #9  Opportunity Subscribed Load
  term                     : String; // FDD #10 Term (months)
  commencementLetterSent   : String; // FDD #11 (LCVP, Sales Admins) "true"/"false"
  commencementLetterSigned : String; // FDD #12 (LCVP, Sales Admins) "true"/"false"
  enrolled                 : String; // FDD #13 free-text from CR&B e.g. "MIGP-FLEX,LCVP"
  meteredCon               : String; // FDD #14 (LCVP)
  unmeteredCon             : String; // FDD #15 (LCVP)
  nte                      : String; // FDD #16 NTE Price
  includeInPipeline        : String; // FDD #18 "true"/"false"
  objectStatus             : String; //To determine data is opportunity or quote.
  salesPhase               : String; //added 20-03-26
  enrollmentReferralCode   : String; // SMB — Referral Code (e.g. TONY), ISU table entry
  enrollmentEmailId        : String; // SMB — Email ID for MIGP welcome letter
  lastFetchedAt            : Timestamp; // persisted timestamp of last CA Fetch/Validate

}
