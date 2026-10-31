/**
 * Menu access positions in USR.USRAccess (VB6 Medix.CheckAccess(n): character n is "1"), from the *_Click handlers of the CIS
 * MDIMain.frm. Mirror of Cis/src/Cis.Api/Auth/AccessRight.cs — names and values must match (AccessRightTests checks it).
 * Forms the desktop never opens are gated like their closest sibling (see the comments in AccessRight.cs).
 */
export const Access = {
  // Membership
  MbmAdjustment: 113,
  MbmEnquiry: 113,
  MbmTopup: 113,
  MbmTracking: 114,
  MbmTrackingJupiter: 160,
  // MCO Fees
  McoInvoice: 115,
  McoFeeListing: 116,
  // Invoice
  InvoiceClinic: 117,
  InvoiceMember: 118,
  InvoiceAdjustment: 119,
  InvoiceAdjustmentAfterBordx: 179,
  EmailNotification: 117,
  GeneralExclusions: 133,
  GeneralExclusionsEnquiry: 113,
  ClaimAscii: 119,
  CaseTransfer: 162,
  InvoiceTracking: 161,
  InvoiceTrackingJupiter: 162,
  MmaCpt: 150,
  WorksheetRegistration: 117,
  WorksheetAdjustment: 119,
  // Claim
  PassToClaim: 163,
  // Claim Bordx
  BordxInsert: 122,
  BordxRemove: 123,
  BordxPrint: 124,
  BordxSendToPayor: 125,
  BordxReturnedFromPayor: 126,
  BordxSendToAc: 127,
  BordxAcPeriod: 128,
  BordxControl: 129,
  BordxOutstanding: 130,
  BordxWorksheet: 131,
  ClaimTracking: 132,
  ClaimTracking2: 120,
  ClaimTrackingJupiter: 163,
  ClaimTracking2Jupiter: 164,
  // LG / Audit
  GenerateLg: 119,
  AuditReview: 119,
  // Payment
  PaymentClinic: 133,
  PaymentClinicInvoice: 121,
  PaymentMember: 134,
  PaymentTracking: 165,
  PaymentTrackingJupiter: 167,
  PaymentFfs: 121,
  // Receipt
  ReceiptMember: 135,
  ReceiptPayor: 136,
  ReceiptFfs: 121,
  // A/C Export
  AcExport: 120,
  // Tracking
  ClinicTracking: 144,
  // Maintenance
  MaintenanceEntity: 137,
  MaintenancePlan: 138,
  MaintenanceCase: 139,
  MaintenanceCalculation: 140,
  MaintenanceLg: 138,
  MaintenanceBenefitMember: 141,
  MaintenanceBenefitClinic: 142,
  MaintenanceCheckItem: 143,
} as const

export type AccessPosition = (typeof Access)[keyof typeof Access]

/** VB6 CheckAccess: true when character `position` (1-based) of the access string is "1". */
export function hasPosition(access: string | null | undefined, position: number): boolean {
  return !!access && position >= 1 && access.length >= position && access.charAt(position - 1) === '1'
}
