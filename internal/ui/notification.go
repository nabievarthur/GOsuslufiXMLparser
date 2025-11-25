package ui

import (
	"image/color"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/canvas"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/layout"
)

type Notifier struct {
	box *fyne.Container
}

func NewNotifier() *Notifier {
	return &Notifier{
		box: container.NewVBox(),
	}
}

func (n *Notifier) Widget() *fyne.Container {
	return n.box
}

func (n *Notifier) Show(message string) {
	notificationText := canvas.NewText(message, color.NRGBA{R: 122, G: 116, B: 119, A: 255})
	notificationText.TextStyle.Bold = true

	notification := container.NewHBox(
		notificationText,
		layout.NewSpacer(),
	)

	n.box.Add(notification)
	n.box.Refresh()

	go func() {
		time.Sleep(3 * time.Second)
		fyne.Do(func() {
			n.box.Remove(notification)
			n.box.Refresh()
		})
	}()
}
