package ui

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
)

func Run() {
	a := app.New()
	w := a.NewWindow("Госуслуги")
	w.Resize(fyne.NewSize(900, 600))
	w.CenterOnScreen()
	ic, _ := fyne.LoadResourceFromPath("../../icon.png")
	w.SetIcon(ic)
	w.SetContent(BuildUI(w))
	w.ShowAndRun()
}
