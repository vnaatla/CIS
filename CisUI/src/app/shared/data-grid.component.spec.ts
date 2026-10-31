import { Component, signal } from '@angular/core'
import { TestBed } from '@angular/core/testing'
import { DataGridComponent, GridColumn } from './data-grid.component'

interface Row {
  code: string
  amount: number
}

@Component({
  imports: [DataGridComponent],
  template: `<app-data-grid caption="Claims" [columns]="columns" [rows]="rows()" [rowKey]="key" [checkable]="true" [checked]="checked()" (checkedChange)="checked.set($event)" />`,
})
class Host {
  columns: GridColumn<Row>[] = [
    { key: 'code', label: 'Code' },
    { key: 'amount', label: 'Amount', align: 'right' },
  ]
  rows = signal<Row[]>([
    { code: 'B', amount: 10 },
    { code: 'A', amount: 2 },
  ])
  checked = signal<Set<string>>(new Set())
  key = (r: Row) => r.code
}

describe('DataGridComponent', () => {
  const cells = (el: HTMLElement) => Array.from(el.querySelectorAll('tbody tr td:nth-child(2)')).map((td) => td.textContent?.trim())

  it('sorts by a column header (ascending, descending, off)', async () => {
    const fixture = TestBed.createComponent(Host)
    await fixture.whenStable()
    const el = fixture.nativeElement as HTMLElement
    expect(cells(el)).toEqual(['B', 'A'])

    const header = el.querySelector<HTMLButtonElement>('th button.th-sort')!
    header.click()
    await fixture.whenStable()
    expect(cells(el)).toEqual(['A', 'B'])
    header.click()
    await fixture.whenStable()
    expect(cells(el)).toEqual(['B', 'A'])
  })

  it('checks all rows from the header box and shows the empty text without rows', async () => {
    const fixture = TestBed.createComponent(Host)
    await fixture.whenStable()
    const el = fixture.nativeElement as HTMLElement
    el.querySelector<HTMLInputElement>('thead input[type=checkbox]')!.click()
    await fixture.whenStable()
    expect([...fixture.componentInstance.checked()].sort()).toEqual(['A', 'B'])

    fixture.componentInstance.rows.set([])
    await fixture.whenStable()
    expect(el.querySelector('tr.empty-row')?.textContent?.trim()).toBe('No records found.')
  })
})
