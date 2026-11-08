using System.Globalization;

namespace Cis.Application.Lg;

/// <summary>What a gate decided: stop (with the desktop's message, if any), ask (OK / Cancel) or carry on.</summary>
public enum LgGateKind
{
    /// <summary>Carry on.</summary>
    Ok,
    /// <summary>Ask the question; on OK call again with one more confirmation.</summary>
    Confirm,
    /// <summary>Stop (Message null = the desktop stopped without a message).</summary>
    Stop,
}

/// <summary>A gate result.</summary>
/// <param name="Kind">Decision.</param>
/// <param name="Message">The question or the stopping message.</param>
/// <param name="Notices">Messages shown on the way that did not stop.</param>
public record LgGateResult(LgGateKind Kind, string? Message, List<string> Notices);

/// <summary>The limit checks of CmdAdd_Click and CmdEdit_Click, translated statement by statement from frmLGGeneration.frm
/// (including the desktop's string concatenation of formatted balances in the "B" branches).</summary>
public static class LgGates
{
    private const string ScomiMsg = "Hospital KPJ Damansara & Prince Court only Applicable For Shah Hakim Zain & Family.(SCHA0000867*01)";

    private sealed class Confirm(string message) : Exception(message);

    private sealed class TypeMismatch : Exception;

    private sealed class Run(int confirmed)
    {
        private int _asked;
        public readonly List<string> Said = [];

        public bool Ask(string message) => _asked++ < confirmed ? true : throw new Confirm(message);
    }

    /// <summary>CmdAdd_Click: may a new LG be started for the clicked cover?</summary>
    public static LgGateResult Add(LgLimits l, bool validLg, bool memberSelected, DateTime? expiry, DateTime today, string group, int confirmed) =>
        Go(confirmed, r => AddTree(r, l, validLg, memberSelected, expiry ?? DateTime.MinValue, today.Date, group));

    /// <summary>CmdEdit_Click's limit checks (after the bordereaux / payment / panel / expiry checks).</summary>
    public static LgGateResult Edit(LgLimits l, string gpSpShared, string group, int confirmed) => Go(confirmed, r => EditTree(r, l, gpSpShared, group));

    private static LgGateResult Go(int confirmed, Func<Run, LgGateKind> tree)
    {
        var run = new Run(confirmed);
        try
        {
            var kind = tree(run);
            return kind == LgGateKind.Stop
                ? new LgGateResult(LgGateKind.Stop, run.Said.Count > 0 ? run.Said[^1] : null, run.Said.Count > 0 ? run.Said[..^1] : [])
                : new LgGateResult(LgGateKind.Ok, null, run.Said);
        }
        catch (Confirm c) { return new LgGateResult(LgGateKind.Confirm, c.Message, run.Said); }
        catch (TypeMismatch) { return new LgGateResult(LgGateKind.Stop, null, run.Said); }
    }

    private static bool IsNum(string v) => LgRules.IsNum(v);

    private static decimal N(string v) => LgRules.IsNum(v) ? LgRules.Num(v) : throw new TypeMismatch();

    /// <summary>Format(CDbl(x)) used as a number.</summary>
    private static decimal G(string v) => N(v);

    private static decimal G(decimal v) => v;

    /// <summary>Format(CDbl(a)) + Format(CDbl(b)): VB joins the two texts (e.g. 100 and 50 give 10050).</summary>
    private static decimal Cat(string a, string b) =>
        decimal.TryParse(Str(N(a)) + Str(N(b)), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : throw new TypeMismatch();

    private static string Str(decimal v) => v.ToString("0.############", CultureInfo.InvariantCulture);

    private static string F(decimal v) => v.ToString("#,##0.00", CultureInfo.InvariantCulture);

    private static LgGateKind AddTree(Run run, LgLimits l, bool validLg, bool memberSelected, DateTime expiry, DateTime today, string group)
    {
        bool Ask(string m) => run.Ask(m);
        void Say(string m) => run.Said.Add(m);
        static LgGateKind Stop() => LgGateKind.Stop;
        decimal cBal = 0;
            if (validLg) {
                if (memberSelected) {
                    if (l.LgIssuedInd == "N") {
                        Say("Letter of Guarantee is Not Allowed !!");
                        return Stop();
                    } else if (expiry < today) {
                        Say("Letter of Guarantee is Not Allowed !!");
                        return Stop();
                    } else if (l.LgIssuedInd == "Y") {
                        if (l.UsingLimit == "A") {
                            if (l.LgIssuedAfterLimitExInd == "N") {
                                if (IsNum(l.SpecBalance)) {
                                    if (IsNum(l.Balance)) {
                                        cBal = N(l.Balance) + N(l.SpecBalance);
                                    }
                                } else {
                                    if (IsNum(l.Balance)) {
                                        cBal = N(l.Balance);
                                    } else {
                                    }
                                }
                                if (IsNum(l.Balance)) {
                                    if (G(cBal) <= 0) {
                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                        return Stop();
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(cBal)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                if (IsNum(l.Balance)) {
                                    if (group == "SCOMI + SUBSIDIARIES") {
                                        if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else {
                                        if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "C") {
                                if (IsNum(l.Balance)) {
                                    if (G(l.Balance) <= 0) {
                                        if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                                return Stop();
                            }
                        } else if (l.UsingLimit == "S") {
                            if (l.LgIssuedAfterLimitExInd == "N") {
                                if (IsNum(l.SpecBalance)) {
                                    if (G(l.SpecBalance) <= 0) {
                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                        return Stop();
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                if (IsNum(l.SpecBalance)) {
                                    if (group == "SCOMI + SUBSIDIARIES") {
                                        if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else {
                                        if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                        Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                        return Stop();
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "C") {
                                if (IsNum(l.SpecBalance)) {
                                    if (G(l.SpecBalance) <= 0) {
                                        if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else if (l.SpecBalance == "FULL") {
                                    if (!Ask("Member Entitled Max LG Amount RM " + l.SpecBalance + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                        Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                        return Stop();
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                                return Stop();
                            }
                        } else if (l.UsingLimit == "B") {
                            if (l.LgIssuedAfterLimitExInd == "N") {
                                if (IsNum(l.Balance)) {
                                    if (IsNum(l.SpecBalance)) {
                                        if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecBalance == "N/A") {
                                            if (l.SpecAnnual == "FULL") {
                                                if (group == "SCOMI + SUBSIDIARIES") {
                                                    if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                } else {
                                                    if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else if (l.SpecAnnual == "N/A") {
                                                if (G(l.Balance) <= 0) {
                                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                    return Stop();
                                                } else {
                                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                return Stop();
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                if (IsNum(l.Balance)) {
                                    if (IsNum(l.SpecBalance)) {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecBalance == "N/A") {
                                            if (l.SpecAnnual == "FULL") {
                                                if (group == "SCOMI + SUBSIDIARIES") {
                                                    if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                } else {
                                                    if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else if (l.SpecAnnual == "N/A") {
                                                if (G(l.Balance) <= 0) {
                                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed.");
                                                    return Stop();
                                                } else {
                                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                return Stop();
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.");
                                        return Stop();
                                    }
                                }
                            } else if (l.LgIssuedAfterLimitExInd == "C") {
                                if (IsNum(l.Balance)) {
                                    if (IsNum(l.SpecBalance)) {
                                        if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                            if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecBalance == "N/A") {
                                            if (l.SpecAnnual == "FULL") {
                                                if (group == "SCOMI + SUBSIDIARIES") {
                                                    if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                } else {
                                                    if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else if (l.SpecAnnual == "N/A") {
                                                if (G(l.Balance) <= 0) {
                                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                    return Stop();
                                                } else {
                                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                        return Stop();
                                                    }
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                return Stop();
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                        return Stop();
                                    }
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                            }
                        } else {
                            Say("Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!");
                        }
                    } else if (l.LgIssuedInd == "C") {
                        if (Ask("Please verify with " + group + " HR Dept or Health Connect ! Press OK to Add LG and Cancel to Abort")) {
                            if (l.UsingLimit == "A") {
                                if (l.LgIssuedAfterLimitExInd == "N") {
                                    if (IsNum(l.Balance)) {
                                        if (G(l.Balance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                    if (IsNum(l.Balance)) {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "C") {
                                    if (IsNum(l.Balance)) {
                                        if (G(l.Balance) <= 0) {
                                            if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                                    return Stop();
                                }
                            } else if (l.UsingLimit == "S") {
                                if (l.LgIssuedAfterLimitExInd == "N") {
                                    if (IsNum(l.SpecBalance)) {
                                        if (G(l.SpecBalance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                    if (IsNum(l.SpecBalance)) {
                                        if (group == "SCOMI + SUBSIDIARIES") {
                                            if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                            Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                            return Stop();
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "C") {
                                    if (IsNum(l.SpecBalance)) {
                                        if (G(l.SpecBalance) <= 0) {
                                            if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                            Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                            return Stop();
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                                    return Stop();
                                }
                            } else if (l.UsingLimit == "B") {
                                if (l.LgIssuedAfterLimitExInd == "N") {
                                    if (IsNum(l.Balance)) {
                                        if (IsNum(l.SpecBalance)) {
                                            if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                                Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                return Stop();
                                            } else {
                                                if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            if (l.SpecBalance == "N/A") {
                                                if (l.SpecAnnual == "FULL") {
                                                    if (group == "SCOMI + SUBSIDIARIES") {
                                                        if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    } else {
                                                        if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else if (l.SpecAnnual == "N/A") {
                                                    if (G(l.Balance) <= 0) {
                                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                        return Stop();
                                                    } else {
                                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else {
                                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                    return Stop();
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "Y") {
                                    if (IsNum(l.Balance)) {
                                        if (IsNum(l.SpecBalance)) {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            if (l.SpecBalance == "N/A") {
                                                if (l.SpecAnnual == "FULL") {
                                                    if (group == "SCOMI + SUBSIDIARIES") {
                                                        if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    } else {
                                                        if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else if (l.SpecAnnual == "N/A") {
                                                    if (G(l.Balance) <= 0) {
                                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed.");
                                                        return Stop();
                                                    } else {
                                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else {
                                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                    return Stop();
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!");
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.");
                                            return Stop();
                                        }
                                    }
                                } else if (l.LgIssuedAfterLimitExInd == "C") {
                                    if (IsNum(l.Balance)) {
                                        if (IsNum(l.SpecBalance)) {
                                            if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                                if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            if (l.SpecBalance == "N/A") {
                                                if (l.SpecAnnual == "FULL") {
                                                    if (group == "SCOMI + SUBSIDIARIES") {
                                                        if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    } else {
                                                        if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else if (l.SpecAnnual == "N/A") {
                                                    if (G(l.Balance) <= 0) {
                                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                        return Stop();
                                                    } else {
                                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Add LG and Cancel to Abort.")) {
                                                            return Stop();
                                                        }
                                                    }
                                                } else {
                                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                                    return Stop();
                                                }
                                            } else {
                                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                            if (group == "SCOMI + SUBSIDIARIES") {
                                                if (!Ask(ScomiMsg + " Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            } else {
                                                if (!Ask("Press OK to Add LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                            return Stop();
                                        }
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!");
                            }
                        } else {
                            return Stop();
                        }
                    } else {
                        Say("Letter of Guarantee is Not Allowed. Please Check LG Issued Ind !!");
                        return Stop();
                    }
                    return LgGateKind.Ok;
                } else {
                    Say("Please select a member before you create the LG");
                }
            } else {
                Say("No LG For this Plan!!!!");
            }
    
        return LgGateKind.Stop;
    }

    private static LgGateKind EditTree(Run run, LgLimits l, string gpSpShared, string group)
    {
        bool Ask(string m) => run.Ask(m);
        void Say(string m) => run.Said.Add(m);
        static LgGateKind Stop() => LgGateKind.Stop;
        decimal chk = 0;
            if (l.LgIssuedInd == "N") {
                Say("Letter of Guarantee is Not Allowed !!");
                return Stop();
            } else if (l.LgIssuedInd == "Y") {
                if (l.UsingLimit == "A") {
                    if (l.LgIssuedAfterLimitExInd == "N") {
                        if (IsNum(l.Balance)) {
                            if (IsNum(l.SpecBalance)) {
                                if ((G(l.Balance) <= 0) && (N(l.SpecBalance) <= 0)) {
                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                    return Stop();
                                } else {
                                    if (gpSpShared == "Y") {
                                        if (G(l.Balance) > 0) {
                                            chk = N(l.Balance);
                                        } else if (G(l.SpecBalance) > 0) {
                                            chk = N(l.SpecBalance);
                                        }
                                    } else {
                                        if (G(l.SpecBalance) > 0) {
                                            chk = N(l.SpecBalance);
                                        }
                                    }
                                    if (chk == 0) {
                                        Say("Insufficient Balance, Please Check");
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(G(chk)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                }
                            } else {
                                if (G(l.Balance) <= 0) {
                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                    return Stop();
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                return Stop();
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "Y") {
                        if (IsNum(l.Balance)) {
                            if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                return Stop();
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                return Stop();
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "C") {
                        if (IsNum(l.Balance)) {
                            if (G(l.Balance) <= 0) {
                                if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!");
                                return Stop();
                            }
                        }
                    } else {
                        Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                        return Stop();
                    }
                } else if (l.UsingLimit == "S") {
                    if (l.LgIssuedAfterLimitExInd == "N") {
                        if (IsNum(l.SpecBalance)) {
                            if (G(l.SpecBalance) <= 0) {
                                Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                return Stop();
                            } else {
                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "Y") {
                        if (IsNum(l.SpecBalance)) {
                            if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                return Stop();
                            }
                        } else {
                            if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                return Stop();
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                return Stop();
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "C") {
                        if (IsNum(l.SpecBalance)) {
                            if (G(l.SpecBalance) <= 0) {
                                if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                return Stop();
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                return Stop();
                            }
                        }
                    } else {
                        Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                        return Stop();
                    }
                } else if (l.UsingLimit == "B") {
                    if (l.LgIssuedAfterLimitExInd == "N") {
                        if (IsNum(l.Balance)) {
                            if (IsNum(l.SpecBalance)) {
                                if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                    return Stop();
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.SpecBalance == "N/A") {
                                    if (l.SpecAnnual == "FULL") {
                                        if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else if (l.SpecAnnual == "N/A") {
                                        if (G(l.Balance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                return Stop();
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "Y") {
                        if (IsNum(l.Balance)) {
                            if (IsNum(l.SpecBalance)) {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                if (l.SpecBalance == "N/A") {
                                    if (l.SpecAnnual == "FULL") {
                                        if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else if (l.SpecAnnual == "N/A") {
                                        if (G(l.Balance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed.");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.");
                                return Stop();
                            }
                        }
                    } else if (l.LgIssuedAfterLimitExInd == "C") {
                        if (IsNum(l.Balance)) {
                            if (IsNum(l.SpecBalance)) {
                                if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                    if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.SpecBalance == "N/A") {
                                    if (l.SpecAnnual == "FULL") {
                                        if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else if (l.SpecAnnual == "N/A") {
                                        if (G(l.Balance) <= 0) {
                                            Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                            return Stop();
                                        } else {
                                            if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                    return Stop();
                                }
                            }
                        } else {
                            if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                return Stop();
                            }
                        }
                    } else {
                        Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                    }
                } else {
                    Say("Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!");
                }
            } else if (l.LgIssuedInd == "C") {
                if (Ask("Please verify with " + group + " HR Dept or Health Connect ! Press OK to Edit LG and Cancel to Abort")) {
                    if (l.UsingLimit == "A") {
                        if (l.LgIssuedAfterLimitExInd == "N") {
                            if (IsNum(l.Balance)) {
                                if (G(l.Balance) <= 0) {
                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                    return Stop();
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "Y") {
                            if (IsNum(l.Balance)) {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "C") {
                            if (IsNum(l.Balance)) {
                                if (G(l.Balance) <= 0) {
                                    if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed . Please Check Annual Limit and Annual Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else {
                            Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                            return Stop();
                        }
                    } else if (l.UsingLimit == "S") {
                        if (l.LgIssuedAfterLimitExInd == "N") {
                            if (IsNum(l.SpecBalance)) {
                                if (G(l.SpecBalance) <= 0) {
                                    Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                    return Stop();
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "Y") {
                            if (IsNum(l.SpecBalance)) {
                                if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                    return Stop();
                                }
                            } else {
                                if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                    Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                    return Stop();
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "C") {
                            if (IsNum(l.SpecBalance)) {
                                if (G(l.SpecBalance) <= 0) {
                                    if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.SpecAnnual == "FULL" && l.SpecBalance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else if (l.SpecAnnual == "N/A" && l.SpecBalance == "N/A") {
                                    Say("Member Not Entitled to Spec Limit. Letter of Guarantee is Not Allowed!!");
                                    return Stop();
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Spec Limit and Spec Balance Limit !!");
                                    return Stop();
                                }
                            }
                        } else {
                            Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                            return Stop();
                        }
                    } else if (l.UsingLimit == "B") {
                        if (l.LgIssuedAfterLimitExInd == "N") {
                            if (IsNum(l.Balance)) {
                                if (IsNum(l.SpecBalance)) {
                                    if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                        Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                        return Stop();
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.SpecBalance == "N/A") {
                                        if (l.SpecAnnual == "FULL") {
                                            if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else if (l.SpecAnnual == "N/A") {
                                            if (G(l.Balance) <= 0) {
                                                Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                return Stop();
                                            } else {
                                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                            return Stop();
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "Y") {
                            if (IsNum(l.Balance)) {
                                if (IsNum(l.SpecBalance)) {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    if (l.SpecBalance == "N/A") {
                                        if (l.SpecAnnual == "FULL") {
                                            if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else if (l.SpecAnnual == "N/A") {
                                            if (G(l.Balance) <= 0) {
                                                Say("Limit Exceeded. Letter of Guarantee is Not Allowed.");
                                                return Stop();
                                            } else {
                                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                            return Stop();
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit !!");
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit !!.");
                                    return Stop();
                                }
                            }
                        } else if (l.LgIssuedAfterLimitExInd == "C") {
                            if (IsNum(l.Balance)) {
                                if (IsNum(l.SpecBalance)) {
                                    if (Cat(l.Balance, l.SpecBalance) <= 0) {
                                        if (!Ask("Limit Exceeded. Please verify with " + group + " HR Dept or Health Connect. Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    } else {
                                        if (!Ask("Member Entitled Max LG Amount RM " + F(Cat(l.Balance, l.SpecBalance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                            return Stop();
                                        }
                                    }
                                } else {
                                    if (l.SpecBalance == "N/A") {
                                        if (l.SpecAnnual == "FULL") {
                                            if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                                return Stop();
                                            }
                                        } else if (l.SpecAnnual == "N/A") {
                                            if (G(l.Balance) <= 0) {
                                                Say("Limit Exceeded. Letter of Guarantee is Not Allowed !!");
                                                return Stop();
                                            } else {
                                                if (!Ask("Member Entitled Max LG Amount RM " + F(G(l.Balance)) + ". Please Notify Member Accordingly. Press OK to Edit LG and Cancel to Abort.")) {
                                                    return Stop();
                                                }
                                            }
                                        } else {
                                            Say("Letter of Guarantee is Not Allowed. Please Check Spec Annual Limit !!");
                                            return Stop();
                                        }
                                    } else {
                                        Say("Letter of Guarantee is Not Allowed. Please Check Spec Balance Limit");
                                        return Stop();
                                    }
                                }
                            } else {
                                if (l.AnnualLimit == "FULL" && l.Balance == "FULL") {
                                    if (!Ask("Press OK to Edit LG and Cancel to Abort.")) {
                                        return Stop();
                                    }
                                } else {
                                    Say("Letter of Guarantee is Not Allowed. Please Check Annual Limit and Annual Balance Limit.");
                                    return Stop();
                                }
                            }
                        } else {
                            Say("Letter of Guarantee is Not Allowed. Please Check LG Issued After Limit Exceeded Ind !!.");
                        }
                    } else {
                        Say("Letter of Guarantee is Not Allowed. Please Check Limit Used for LG !!");
                    }
                } else {
                    return Stop();
                }
            } else {
                Say("Letter of Guarantee is Not Allowed. Please Check LG Issued Ind !!");
                return Stop();
            }
            return LgGateKind.Ok;
    
    }
}
