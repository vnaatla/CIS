// Generates Cis/sql/main/1140_mbm_tracking.sql: cis.MbmTrk_Run, the cases of HISDB CISMBMTracking with the pasted criteria replaced by a parameterized WHERE fragment.
const fs = require('fs')
const src = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/CISMBMTracking.sql', 'utf8').replace(/\r\n/g, '\n')
const parts = src.split(/\n\s*if @strAppendCase = (\d+)\s*\n/)
const cases = []
for (let i = 1; i < parts.length; i += 2) {
  const n = Number(parts[i]); const body = parts[i + 1]
  // String literal assignments: set @var = N'...' or '...'
  const lits = {}
  const re = /set\s+(@\w+)\s*=\s*N?'((?:[^']|'')*)'/gi
  let m
  while ((m = re.exec(body))) lits[m[1].toLowerCase()] = m[2]
  const comp = /set\s+(@\w+)\s*=\s*(@\w+)\s*\+\s*@strAppend(?:\s*\+\s*(@\w+))?/i.exec(body)
  if (!comp) throw new Error('case ' + n + ': no composition')
  const before = lits[comp[2].toLowerCase()]
  const after = comp[3] ? lits[comp[3].toLowerCase()] : ''
  if (before === undefined || after === undefined) throw new Error('case ' + n + ': literal missing')
  cases.push({ n, before, after })
}
const esc = (s) => s.replace(/'/g, "''") // the literals are already in doubled-quote form; re-emit as they were
const filters = fs.readFileSync('C:/Users/Naresh/AppData/Local/Temp/mbmtrk-filters.sql', 'utf8')
let out = `-- Tracking > MBM Tracking (FrmNewMBMTracking and its variants): the ${cases.length} cases of HISDB CISMBMTracking (generated from the procedure by
-- gen-mbmtrk.js). Each case's query text is kept as the desktop's; the pasted criteria are replaced by a WHERE fragment built only from the filters that are set,
-- with every value passed as a parameter.
SET QUOTED_IDENTIFIER ON;
GO

CREATE OR ALTER PROCEDURE cis.MbmTrk_Run @case int,
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
    EXEC HISDB.sys.sp_executesql @sql, N'${filters.split('\n').filter((l) => l.startsWith('--PARAM ')).map((l) => l.slice(8).replace(/\s*=\s*NULL$|\s*=\s*0$/, '')).join(', ')}',
        ${filters.split('\n').filter((l) => l.startsWith('--PARAM ')).map((l) => l.slice(8).split(' ')[0]).join(', ')};
END
GO
`
fs.writeFileSync('D:/CIS/Cis/sql/main/1140_mbm_tracking.sql', out)
console.log('cases', cases.length, cases.slice(0, 3).map((c) => c.n).join(','), '...', cases[cases.length - 1].n)
