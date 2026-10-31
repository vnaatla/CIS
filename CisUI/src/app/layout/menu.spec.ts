import { Access, hasPosition } from '../core/auth/access'
import { MENU, buildable } from './menu'

const items = MENU.flatMap((m) => m.items.map((i) => ({ module: m, item: i })))

describe('access (VB6 CheckAccess)', () => {
  it('reads character n of USRAccess', () => {
    expect(hasPosition('101', 1)).toBe(true)
    expect(hasPosition('101', 2)).toBe(false)
    expect(hasPosition('101', 4)).toBe(false)
    expect(hasPosition(null, 1)).toBe(false)
    expect(hasPosition('1', 0)).toBe(false)
  })
})

describe('CIS menu (MDIMain.frm)', () => {
  it('has a unique route for every item, inside its module', () => {
    const paths = items.map((x) => x.item.path)
    expect(new Set(paths).size).toBe(paths.length)
    for (const { module, item } of items) expect(item.path.startsWith(module.path + '/')).toBe(true)
  })

  it('gates every item with a position from the access table', () => {
    const positions = new Set<number>(Object.values(Access))
    for (const { item } of items) expect(positions.has(item.access)).toBe(true)
  })

  it('names the VB6 form of every item that needs a screen', () => {
    for (const m of MENU) for (const item of buildable(m)) expect(item.form).toMatch(/^(frm|Frm)/)
  })

  it('covers every form the desktop project opens from its menu', () => {
    // A sample across all top menus (MDIMain *_Click handlers); docs/menu-mapping.md has the full list.
    const forms = items.flatMap((x) => x.item.form.split(',').map((f) => f.trim()))
    for (const form of [
      'FrmMBMCLNAdj', 'frmMCOFeeGSTZeroised', 'frmInvoiceAdjAftBordx', 'frmClaimASCIICIMB', 'frmPass2Claim', 'frmPrintBordxBSI', 'FrmSendtoAC',
      'frmNewClaimRpt', 'frmLGGeneration', 'FrmAuditInvoiceTracking', 'frmPaymentMemberAuto08', 'frmReceiptUpMultiple', 'frmPostingInvClnGST2017',
      'frmSUNSystemAllocation', 'FrmTracking2', 'frmMaintenanceBNFDental', 'frmCheckingItems', 'frmTest',
    ])
      expect(forms).toContain(form)
  })

  it('lists the desktop top menus in desktop order (LG + Audit review share one module; File / Windows / Exit live in the top bar)', () => {
    expect(MENU.map((m) => m.label)).toEqual([
      'Membership', 'MCO Fees', 'Invoice', 'Claim', 'Claim Bordx', 'LG & Audit', 'Payment', 'Receipt', 'A/C Export', 'Tracking', 'Maintenance',
    ])
  })
})
