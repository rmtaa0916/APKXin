// RETRI QC 100 — VISUAL HP + FAST + GATE + OCR + LEVEL/DAMAGE
// Diagnostic only. NO RETRIBUTION TAPS. NO GAMEPLAY CLICKS.
// Run inside the same Macrorify project so boss_health / boss_health2 assets are available.
//
// HUD protocol:
//   001-020 BASELINE: objective visible, avoid dummy/AoE pressure.
//   021-080 PRESSURE: add dummy bots and use AoE so the top HP overlay switches often.
//   081-100 EXEC RANGE: keep pressure, lower Turtle/Lord near/below RETRI execute HP.
// When complete, the full report is copied automatically to Clipboard.

var QC_TOTAL = 100
var QC_SAMPLE_WAIT_MS = 250
var QC_TOTAL_TOL = 250
var QC_REF_W = 2340
var QC_REF_H = 1080

var qcDevice = Region.deviceReg()
var qcDeviceW = qcDevice.getW()
var qcDeviceH = qcDevice.getH()
var qcScale = qcDeviceH / QC_REF_H
var qcExtraX = qcDeviceW - QC_REF_W * qcScale

fun qcX(v) {
    return Math.round(v * qcScale + qcExtraX / 2)
}
fun qcY(v) {
    return Math.round(v * qcScale)
}
fun qcS(v) {
    return Math.round(v * qcScale)
}
fun qcAbs(v) {
    if (v < 0) { return 0 - v }
    return v
}
fun qcPct(num, den) {
    if (den <= 0) { return -1 }
    return Math.round((num * 1000) / den) / 10
}
fun qcObjectiveName(t) {
    if (t == 1) { return "T" }
    if (t == 2) { return "L" }
    return "-"
}
fun qcPhase(n) {
    if (n <= 20) { return "BASELINE" }
    if (n <= 80) { return "PRESSURE" }
    return "EXEC-RANGE"
}
fun qcPhaseInstruction(n) {
    if (n <= 20) {
        return "KEEP TURTLE/LORD VISIBLE | NO DUMMY/AOE PRESSURE"
    }
    if (n <= 80) {
        return "ADD DUMMY BOTS + AOE | FORCE TOP HP TO SWITCH"
    }
    return "KEEP PRESSURE | LOWER OBJECTIVE NEAR/BELOW EXEC | DO NOT PRESS RETRI"
}

// Current reference regions, adapted to the device using the production 2340x1080 frame.
var qcTopHpReg = Region(qcX(1015), qcY(56), qcS(309), qcS(60)).noScale()
var qcLevelReg = Region(qcX(1023), qcY(343), qcS(105), qcS(101)).noScale()
var qcDamageReg = Region(qcX(1400), qcY(850), qcS(180), qcS(60)).noScale()

var qcPairParam = TParam.scale(2)
qcPairParam.mode(1)
qcPairParam.whitelist("0123456789/")

var qcLevelParam = TParam.scale(2)
qcLevelParam.mode(1)
qcLevelParam.whitelist("0123456789")

var qcDamageParam = TParam.scale(2)
qcDamageParam.mode(1)
qcDamageParam.whitelist("0123456789")

var qcHud = OnScreenText(qcX(610), qcY(735), qcS(1550), qcS(250), true).noScale()
qcHud.setBackgroundColor(null)
qcHud.setTextColor("#FFFFFF00")
qcHud.setTextSize(11)
qcHud.clickable(false)
qcHud.hidden(false)
qcHud.show()

var qcBarLeft = qcX(1091)
var qcBarAnchor = qcX(1094)
var qcBarWidthRef = 220

fun qcIsRed(c) {
    if (c == null) { return 0 }
    var r = c.red()
    if (r < 100) { return 0 }
    var g = c.green()
    if (g > 110) { return 0 }
    var b = c.blue()
    if (b > 110) { return 0 }
    if (r - g < 40) { return 0 }
    if (r - b < 30) { return 0 }
    return 1
}

fun qcDamageToLevel(dmg) {
    if (dmg == 900) { return 1 }
    if (dmg == 1050) { return 2 }
    if (dmg == 1200) { return 3 }
    if (dmg == 1350) { return 4 }
    if (dmg == 1500) { return 5 }
    if (dmg == 1650) { return 6 }
    if (dmg == 1800) { return 7 }
    if (dmg == 1950) { return 8 }
    if (dmg == 2100) { return 9 }
    if (dmg == 2250) { return 10 }
    if (dmg == 2400) { return 11 }
    if (dmg == 2550) { return 12 }
    if (dmg == 2700) { return 13 }
    if (dmg == 2850) { return 14 }
    if (dmg == 3000) { return 15 }
    return -1
}

fun qcReadPairCached() {
    var out = [-1, -1]
    var raw = qcTopHpReg.readAsString(qcPairParam, "")
    if (raw == "") { return out }
    var parts = raw.split("/")
    if (parts.size != 2) { return out }
    var hp = Num.parse(parts[0], -1)
    var total = Num.parse(parts[1], -1)
    if (hp < 1) { return out }
    if (total < 1000) { return out }
    if (total > 999999) { return out }
    if (hp > total) { return out }
    out[0] = hp
    out[1] = total
    return out
}

fun qcGateCached() {
    var t = qcTopHpReg.find("boss_health", 2)
    if (t != null) { return 1 }
    var l = qcTopHpReg.find("boss_health2", 2)
    if (l != null) { return 2 }
    return 0
}

fun qcReadRowVisualHpCached(totalHp, rowRefY) {
    if (totalHp < 1000) { return -1 }
    var pts = [0]
    pts.clear()
    for (var off = 2; off < qcBarWidthRef - 2; off = off + 3) {
        pts.push(Point(qcBarLeft + qcS(off), qcY(rowRefY)).noScale())
    }
    var colors = Color.getAll(pts)
    if (colors == null) { return -1 }
    var lastRed = -1
    var seenRed = 0
    var emptyRun = 0
    for (var i = 0; i < colors.size; i = i + 1) {
        if (qcIsRed(colors[i]) == 1) {
            seenRed = 1
            lastRed = i
            emptyRun = 0
        } else {
            if (seenRed == 1) {
                emptyRun = emptyRun + 1
                if (emptyRun >= 3) { break }
            }
        }
    }
    colors = null
    if (lastRed < 0) { return -1 }
    var filledRefPx = 2 + (lastRed * 3) + 3
    if (filledRefPx < 1) { filledRefPx = 1 }
    if (filledRefPx > qcBarWidthRef) { filledRefPx = qcBarWidthRef }
    var hp = Math.round((filledRefPx * totalHp) / qcBarWidthRef)
    if (hp < 1) { hp = 1 }
    if (hp > totalHp) { hp = totalHp }
    return hp
}

fun qcMedian3(a, b, c) {
    var lo = a
    var hi = a
    if (b < lo) { lo = b }
    if (c < lo) { lo = c }
    if (b > hi) { hi = b }
    if (c > hi) { hi = c }
    return a + b + c - lo - hi
}

fun qcVisualHpCached(totalHp) {
    var out = [-1, -1, -1, -1]
    var a = qcReadRowVisualHpCached(totalHp, 84)
    var b = qcReadRowVisualHpCached(totalHp, 87)
    var c = qcReadRowVisualHpCached(totalHp, 90)
    out[1] = a
    out[2] = b
    out[3] = c
    if (a < 1) { return out }
    if (b < 1) { return out }
    if (c < 1) { return out }
    out[0] = qcMedian3(a, b, c)
    return out
}

// Reproduce current R67/R62 FAST classifier on the same cached frame.
// Return: 1=FIRE, 0=BLOCK/HIGH, 2=INVALID.
fun qcFastClassifierCached(execDmg, totalHp) {
    if (execDmg < 1) { return 2 }
    if (totalHp < 1000) { return 2 }
    if (execDmg >= totalHp) { return 2 }

    var offsetX100 = (execDmg * 100 * qcBarWidthRef) / totalHp
    var offsetRef = 0
    for (var k = 1; k < qcBarWidthRef; k = k + 1) {
        // Current R67 nearest-pixel alignment.
        if (k * 100 <= offsetX100 + 50) { offsetRef = k }
    }
    if (offsetRef < 6) { return 2 }
    if (offsetRef > qcBarWidthRef - 4) { return 2 }

    var thresholdX = qcBarLeft + qcS(offsetRef)
    var guardX = thresholdX - qcS(7)
    var out1 = thresholdX + qcS(3)
    var out2 = thresholdX + qcS(6)
    var out3 = thresholdX + qcS(9)

    var pts = [0]
    pts.clear()
    pts.push(Point(qcBarAnchor, qcY(82)).noScale())
    pts.push(Point(qcBarAnchor, qcY(87)).noScale())
    pts.push(Point(qcBarAnchor, qcY(92)).noScale())
    pts.push(Point(guardX, qcY(82)).noScale())
    pts.push(Point(guardX, qcY(87)).noScale())
    pts.push(Point(guardX, qcY(92)).noScale())
    pts.push(Point(thresholdX, qcY(82)).noScale())
    pts.push(Point(thresholdX, qcY(87)).noScale())
    pts.push(Point(thresholdX, qcY(92)).noScale())
    pts.push(Point(thresholdX - qcS(1), qcY(82)).noScale())
    pts.push(Point(thresholdX - qcS(1), qcY(87)).noScale())
    pts.push(Point(thresholdX - qcS(1), qcY(92)).noScale())
    pts.push(Point(thresholdX + qcS(1), qcY(82)).noScale())
    pts.push(Point(thresholdX + qcS(1), qcY(87)).noScale())
    pts.push(Point(thresholdX + qcS(1), qcY(92)).noScale())
    pts.push(Point(out1, qcY(82)).noScale())
    pts.push(Point(out1, qcY(87)).noScale())
    pts.push(Point(out1, qcY(92)).noScale())
    pts.push(Point(out2, qcY(82)).noScale())
    pts.push(Point(out2, qcY(87)).noScale())
    pts.push(Point(out2, qcY(92)).noScale())
    pts.push(Point(out3, qcY(82)).noScale())
    pts.push(Point(out3, qcY(87)).noScale())
    pts.push(Point(out3, qcY(92)).noScale())

    var colors = Color.getAll(pts)
    if (colors == null) { return 2 }
    if (colors.size < 24) {
        colors = null
        return 2
    }

    var anchorRed = 0
    var guardRed = 0
    var thresholdRed = 0
    if (qcIsRed(colors[0]) == 1) { anchorRed = anchorRed + 1 }
    if (qcIsRed(colors[1]) == 1) { anchorRed = anchorRed + 1 }
    if (qcIsRed(colors[2]) == 1) { anchorRed = anchorRed + 1 }
    if (qcIsRed(colors[3]) == 1) { guardRed = guardRed + 1 }
    if (qcIsRed(colors[4]) == 1) { guardRed = guardRed + 1 }
    if (qcIsRed(colors[5]) == 1) { guardRed = guardRed + 1 }
    if (qcIsRed(colors[6]) == 1) { thresholdRed = thresholdRed + 1 }
    if (qcIsRed(colors[7]) == 1) { thresholdRed = thresholdRed + 1 }
    if (qcIsRed(colors[8]) == 1) { thresholdRed = thresholdRed + 1 }

    if (thresholdRed >= 2) {
        colors = null
        return 0
    }
    if (guardRed < 2) {
        colors = null
        return 2
    }
    if (anchorRed < 1) {
        colors = null
        return 2
    }

    var confirmRed = 0
    for (var j = 9; j < 15; j = j + 1) {
        if (qcIsRed(colors[j]) == 1) { confirmRed = confirmRed + 1 }
    }
    if (confirmRed >= 4) {
        colors = null
        return 2
    }

    var emptyColumns = 0
    var outwardRedTotal = 0
    for (var col = 0; col < 3; col = col + 1) {
        var columnRed = 0
        var pointStart = 15 + col * 3
        for (var row = 0; row < 3; row = row + 1) {
            if (qcIsRed(colors[pointStart + row]) == 1) {
                columnRed = columnRed + 1
                outwardRedTotal = outwardRedTotal + 1
            }
        }
        if (columnRed <= 1) { emptyColumns = emptyColumns + 1 }
    }
    colors = null

    if (emptyColumns < 2) { return 0 }
    if (outwardRedTotal > 3) { return 0 }
    return 1
}

fun qcFastText(v) {
    if (v == 1) { return "FIRE" }
    if (v == 0) { return "BLOCK" }
    return "INVALID"
}

fun qcSafeText(v) {
    if (v < 0) { return "-" }
    return "" + v
}

fun runRetributionQc100() {
    var validVhp = 0
    var invalidVhp = 0
    var ocrValid = 0
    var ocrInvalid = 0
    var gateOk = 0
    var gateMiss = 0
    var gateSwitch = 0
    var totalStable = 0
    var totalUnstable = 0
    var compareCount = 0
    var errSum = 0
    var errMax = 0
    var errOver5 = 0
    var errOver10 = 0
    var fastTP = 0
    var fastTN = 0
    var fastFP = 0
    var fastFN = 0
    var fastInvalid = 0
    var lvlDmgOk = 0
    var lvlDmgBad = 0
    var lvlDmgUnread = 0
    var baselineType = 0
    var baselineTotal = -1
    var lastDamage = -1
    var lastLevel = -1
    var reportRows = ""

    qcHud.setText("<b>RETRI QC 100 | DIAGNOSTIC ONLY | NO RETRI TAPS\nSTARTING IN 4s | PLACE TURTLE/LORD ON SCREEN</b>")
    wait(4000)

    for (var n = 1; n <= QC_TOTAL; n = n + 1) {
        var phase = qcPhase(n)
        var instruction = qcPhaseInstruction(n)

        Cache.screen()
        var gate = qcGateCached()
        var pair = qcReadPairCached()
        var hp = pair[0]
        var total = pair[1]

        if (baselineType == 0) {
            if (gate > 0) {
                if (total >= 1000) {
                    baselineType = gate
                    baselineTotal = total
                }
            }
        }

        var visualTotal = total
        if (baselineTotal >= 1000) { visualTotal = baselineTotal }
        var v = qcVisualHpCached(visualTotal)
        var vhp = v[0]
        var fastState = 2
        if (lastDamage >= 1) {
            if (visualTotal >= 1000) {
                fastState = qcFastClassifierCached(lastDamage, visualTotal)
            }
        }

        // Level/damage cross-QC every 5th sample to limit OCR load.
        if (n == 1 || n % 5 == 0) {
            var levelRaw = qcLevelReg.readAsString(qcLevelParam, "")
            var dmgRaw = qcDamageReg.readAsString(qcDamageParam, "")
            var levelNow = Num.parse(levelRaw, -1)
            var dmgNow = Num.parse(dmgRaw, -1)
            if (levelNow >= 1) {
                if (levelNow <= 15) { lastLevel = levelNow }
            }
            if (dmgNow >= 900) {
                if (dmgNow <= 3000) { lastDamage = dmgNow }
            }
            var dmgLevel = qcDamageToLevel(lastDamage)
            if (lastLevel < 1) {
                lvlDmgUnread = lvlDmgUnread + 1
            } else if (dmgLevel < 1) {
                lvlDmgUnread = lvlDmgUnread + 1
            } else if (dmgLevel == lastLevel) {
                lvlDmgOk = lvlDmgOk + 1
            } else {
                lvlDmgBad = lvlDmgBad + 1
            }
        }
        Cache.screenOff()

        if (gate == 0) {
            gateMiss = gateMiss + 1
        } else if (baselineType == 0) {
            gateOk = gateOk + 1
        } else if (gate == baselineType) {
            gateOk = gateOk + 1
        } else {
            gateSwitch = gateSwitch + 1
        }

        if (hp >= 1 && total >= 1000) {
            ocrValid = ocrValid + 1
        } else {
            ocrInvalid = ocrInvalid + 1
        }

        if (vhp >= 1) {
            validVhp = validVhp + 1
        } else {
            invalidVhp = invalidVhp + 1
        }

        var totalOk = 0
        if (baselineTotal >= 1000) {
            if (total >= baselineTotal - QC_TOTAL_TOL) {
                if (total <= baselineTotal + QC_TOTAL_TOL) { totalOk = 1 }
            }
        }
        if (totalOk == 1) {
            totalStable = totalStable + 1
        } else {
            if (total >= 1000) { totalUnstable = totalUnstable + 1 }
        }

        var err = -1
        var errPct = -1
        // Only score VHP accuracy when the same objective gate and total are stable.
        if (hp >= 1) {
            if (totalOk == 1) {
                if (gate == baselineType) {
                    if (vhp >= 1) {
                        err = qcAbs(vhp - hp)
                        errPct = qcPct(err, baselineTotal)
                        compareCount = compareCount + 1
                        errSum = errSum + err
                        if (err > errMax) { errMax = err }
                        if (errPct > 5) { errOver5 = errOver5 + 1 }
                        if (errPct > 10) { errOver10 = errOver10 + 1 }
                    }
                }
            }
        }

        // FAST confusion matrix only when OCR benchmark is same-objective + stable-total.
        if (hp >= 1) {
            if (totalOk == 1) {
                if (gate == baselineType) {
                    if (lastDamage >= 1) {
                        var expectedFire = 0
                        if (hp <= lastDamage) { expectedFire = 1 }
                        if (fastState == 2) {
                            fastInvalid = fastInvalid + 1
                        } else if (expectedFire == 1) {
                            if (fastState == 1) { fastTP = fastTP + 1 } else { fastFN = fastFN + 1 }
                        } else {
                            if (fastState == 1) { fastFP = fastFP + 1 } else { fastTN = fastTN + 1 }
                        }
                    }
                }
            }
        }

        var line = "#" + n +
            " P=" + phase +
            " G=" + qcObjectiveName(gate) +
            " OCR=" + qcSafeText(hp) + "/" + qcSafeText(total) +
            " VHP=" + qcSafeText(vhp) +
            " ROW=" + qcSafeText(v[1]) + "," + qcSafeText(v[2]) + "," + qcSafeText(v[3]) +
            " ERR=" + qcSafeText(err) +
            " ERRP=" + qcSafeText(errPct) +
            " EXEC=" + qcSafeText(lastDamage) +
            " LVL=" + qcSafeText(lastLevel) +
            " FAST=" + qcFastText(fastState)
        reportRows = reportRows + line + "\n"

        var hud = "<b>RETRI QC " + n + "/100 | " + phase + " | NO RETRI TAP\n" +
            instruction + "\n" +
            "GATE " + qcObjectiveName(gate) +
            " | OCR " + qcSafeText(hp) + "/" + qcSafeText(total) +
            " | VHP " + qcSafeText(vhp) +
            " | ERR " + qcSafeText(err) + " (" + qcSafeText(errPct) + "%)\n" +
            "EXEC " + qcSafeText(lastDamage) +
            " | LVL " + qcSafeText(lastLevel) +
            " | FAST " + qcFastText(fastState) +
            " | BASE " + qcObjectiveName(baselineType) + "/" + qcSafeText(baselineTotal) + "</b>"
        qcHud.setText(hud)

        wait(QC_SAMPLE_WAIT_MS)
    }

    var meanErr = -1
    if (compareCount > 0) { meanErr = Math.round(errSum / compareCount) }

    var summary =
        "RETRI QC100 R67 VISUAL-HP ISOLATION\n" +
        "BASE=" + qcObjectiveName(baselineType) + "/" + qcSafeText(baselineTotal) + "\n" +
        "VHP valid=" + validVhp + " invalid=" + invalidVhp + "\n" +
        "OCR valid=" + ocrValid + " invalid=" + ocrInvalid + "\n" +
        "GATE same=" + gateOk + " miss=" + gateMiss + " switch=" + gateSwitch + "\n" +
        "TOTAL stable=" + totalStable + " unstable=" + totalUnstable + "\n" +
        "VHPvsOCR compared=" + compareCount + " meanAbsErr=" + meanErr +
        " maxErr=" + errMax + " >5%=" + errOver5 + " >10%=" + errOver10 + "\n" +
        "FASTvsOCR TP=" + fastTP + " TN=" + fastTN + " FP=" + fastFP +
        " FN=" + fastFN + " INVALID=" + fastInvalid + "\n" +
        "LEVELvsDMG ok=" + lvlDmgOk + " bad=" + lvlDmgBad + " unread=" + lvlDmgUnread + "\n" +
        "---SAMPLES---\n" + reportRows

    Clipboard.copy(summary)
    Con.out(summary)

    qcHud.setText("<b>RETRI QC 100 COMPLETE | REPORT COPIED\nPASTE CLIPBOARD BACK TO CHATGPT\n" +
        "VHP " + validVhp + "/100 | OCR " + ocrValid + "/100 | GATE MISS " + gateMiss +
        " | FAST FP " + fastFP + " | FAST FN " + fastFN +
        " | LEVEL BAD " + lvlDmgBad + "</b>")
    qcHud.setTextColor("#FF34D399")
}

runRetributionQc100()