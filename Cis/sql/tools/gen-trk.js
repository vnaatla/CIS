// Generates a tracking query engine: the cases of a desktop tracking procedure with the pasted criteria replaced by a parameterized WHERE fragment.
// Usage: node gen-trk.js mbm | claim
const fs = require('fs')
const T = 'C:/Users/Naresh/AppData/Local/Temp/'
const CFG = {
  mbm: { src: T + 'CISMBMTracking.sql', filters: T + 'mbmtrk-filters.sql', out: 'D:/CIS/Cis/sql/main/1140_mbm_tracking.sql', proc: 'cis.MbmTrk_Run', db: 'HISDB',
    title: 'Tracking > MBM Tracking (FrmNewMBMTracking and its variants): the {n} cases of HISDB CISMBMTracking' },
  mbmsimple: { src: T + 'CISMBMTracking_Simple.sql', filters: T + 'mbmtrk-filters.sql', out: 'D:/CIS/Cis/sql/main/1142_mbm_tracking_simple.sql', proc: 'cis.MbmTrkSimple_Run', db: 'HISDB',
    title: 'Tracking > MBM Tracking (Simple) (frmNewMBMTrackingSimple and its Jupiter copy): the {n} cases of HISDB CISMBMTracking_Simple' },
  claim: { src: T + 'CISClaimTracking.sql', filters: T + 'clmtrk-filters.sql', out: 'D:/CIS/Cis/sql/main/1150_claim_tracking.sql', proc: 'cis.ClmTrk_Run', db: 'CISWorksheet',
    title: 'Claim / Bordx > Claim Tracking (frmClaimRpt and its variants): the {n} cases of CISWorksheet CISClaimTracking' },
}[process.argv[2]]
if (!CFG) throw new Error('usage: node gen-trk.js mbm | mbmsimple | claim')
const src = fs.readFileSync(CFG.src, 'utf8').replace(/\r\n/g, '\n')
const parts = src.split(/\n\s*if @strAppendCase = (\d+)\s*\n/)
const cases = []
const extra = []
const skipped = []
for (let i = 1; i < parts.length; i += 2) {
  const n = Number(parts[i]); const body = parts[i + 1]
  // String literal assignments: set @var = N'...' or '...'
  const lits = {}
  const re = /set\s+(@\w+)\s*=\s*N?'((?:[^']|'')*)'/gi
  let m
  while ((m = re.exec(body))) lits[m[1].toLowerCase()] = m[2]
  const comp = /set\s+(@\w+)\s*=\s*(@\w+)\s*\+\s*@strAppend(?:\s*\+\s*(@\w+))?/i.exec(body)
  // A case without pasted criteria (claim case 10: a treatment reason lookup) is not part of the engine.
  if (!comp) { skipped.push(n); continue }
  const before = lits[comp[2].toLowerCase()]
  let after = comp[3] ? lits[comp[3].toLowerCase()] : ''
  // Pieces appended after the criteria (e.g. case 178's pivot): set @sql = @sql + @piece.
  const more = new RegExp('set\\s+' + comp[1] + '\\s*=\\s*' + comp[1] + '\\s*\\+\\s*(@\\w+)', 'gi')
  more.lastIndex = comp.index + comp[0].length
  let mm
  while ((mm = more.exec(body))) { if (lits[mm[1].toLowerCase()] === undefined) throw new Error('case ' + n + ': piece missing'); after += lits[mm[1].toLowerCase()]; extra.push(n) }
  if (before === undefined || after === undefined) throw new Error('case ' + n + ': literal missing')
  // The desktop's procedure runs with QUOTED_IDENTIFIER OFF and some cases quote values with double quotes; written as single-quoted values here.
  const dq = (t) => t.replace(/"([^"]*)"/g, "''$1''")
  cases.push({ n, before: dq(before), after: dq(after) })
}
const esc = (s) => s.replace(/'/g, "''") // the literals are already in doubled-quote form; re-emit as they were
const filters = fs.readFileSync(CFG.filters, 'utf8')
let out = `-- ${CFG.title.replace('{n}', cases.length)} (generated from the procedure by
-- gen-trk.js). Each case's query text is kept as the desktop's; the pasted criteria are replaced by a WHERE fragment built only from the filters that are set,
-- with every value passed as a parameter.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE ${CFG.proc} @case int,
${filters.split('\n').filter((l) => l.startsWith('--PARAM ')).map((l) => '    ' + l.slice(8)).join(',\n')}
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @b nvarchar(max), @a nvarchar(max) = N'';
${cases.map((c) => `    IF @case = ${c.n} SELECT @b = N'${c.before}', @a = N'${c.after}';`).join('\n')}
    IF @b IS NULL RETURN;
    DECLARE @w nvarchar(max) = N'';
${filters.split('\n').filter((l) => l.startsWith('--WHERE ')).map((l) => '    ' + l.slice(8)).join('\n')}
    DECLARE @sql nvarchar(max) = CAST(@b AS nvarchar(max)) + @w + N' ' + @a;
    EXEC ${CFG.db}.sys.sp_executesql @sql, N'${filters.split('\n').filter((l) => l.startsWith('--PARAM ')).map((l) => l.slice(8).replace(/\s*=\s*NULL$|\s*=\s*0$/, '')).join(', ')}',
        ${filters.split('\n').filter((l) => l.startsWith('--PARAM ')).map((l) => l.slice(8).split(' ')[0]).join(', ')};
END
GO
`
fs.writeFileSync(CFG.out, out)
console.log('extra pieces in cases', extra.join(','), '; skipped', skipped.join(','));
console.log('cases', cases.length, cases.slice(0, 3).map((c) => c.n).join(','), '...', cases[cases.length - 1].n)
