#let orange = rgb("#D96422")
#let green = rgb("#6E8C2B")

#let script-size = 7.97224pt

#let conf(doc) = {
  set text(font: "Segoe UI")
  set heading(numbering: "1.")
  show heading: it => {
    if it.level == 1 {
      set text(size: 18pt, fill: orange, weight: 700)
      it
    }
    if it.level == 2 {
      set text(size: 16pt, fill: green, weight: 400)
      it
    }
    if it.level == 3 {
      set text(size: 12pt, fill: green, weight: 700)
      it
    }
  }
  doc
}
