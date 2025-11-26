package ui

import (
	"os"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"
	"github.com/sqweek/dialog"

	"nabievarthur/GOsuslugiXML/internal/service"
)

func OpenFileDialog() string {
	fileName, err := dialog.File().
		Filter("Xml files", "xml").
		Title("Выберите файл XML").
		Load()
	if err != nil {
		return ""
	}
	return fileName
}

func OpenFileDialog1() string {
	fileName, err := dialog.File().
		Filter("xls files", "xls").
		Title("Выберите файл xls").
		Load()
	if err != nil {
		return ""
	}
	return fileName
}

// Функция для создания содержимого вкладки аккордеона с кнопкой копирования
func createTabContent(text string, win fyne.Window, notifier *Notifier, xmlFileName string) fyne.CanvasObject {
	
	entry := widget.NewMultiLineEntry()
	entry.SetText(text)
	entry.SetMinRowsVisible(4)
	entry.Wrapping = fyne.TextWrapWord

	var copyBtn *widget.Button
	var mergeBtn *widget.Button

	mergeBtn = widget.NewButtonWithIcon("Сравнить c ИБД-Ф", theme.SearchReplaceIcon(), func() {
		notifier.Show("Сравнение и создание нового файла...")

		go func() {
			xmlFile := xmlFileName
			xlsFile := OpenFileDialog1()
			if xlsFile == "" {
				fyne.Do(func() {
					notifier.Show("Файл XLS не выбран")
				})
				return
			}

			// Чтение XML и XLS
			xmlData, err := os.ReadFile(xmlFile)
			if err != nil {
				fyne.Do(func() {
					notifier.Show("Ошибка чтения XML: " + err.Error())
				})
				return
			}

			xlsRows, err := service.ReadXLSFile(xlsFile)
			if err != nil {
				fyne.Do(func() {
					notifier.Show("Ошибка чтения XLS: " + err.Error())
				})
				return
			}

			// Сравнение
			matchedRows, err := service.MatchXMLWithXLS(xmlData, xlsRows)
			if err != nil {
				fyne.Do(func() {
					notifier.Show("Ошибка сравнения: " + err.Error())
				})
				return
			}

			// Мутим новый файл
			err = service.ModifyXLSFile(xlsFile, matchedRows)
			if err != nil {
				fyne.Do(func() {
					notifier.Show("Ошибка создания файла: " + err.Error())
				})
				return
			}

			fyne.Do(func() {
				notifier.Show("Новый файл успешно создан!")
			})

		}()
	})
	mergeBtn.Importance = widget.HighImportance
	mergeBtn.Hide()

	// кнопка копирования для этой вкладки
	copyBtn = widget.NewButtonWithIcon("Копировать текст в буфер обмена", theme.ContentCopyIcon(), func() {
		win.Clipboard().SetContent(entry.Text)
		notifier.Show("Текст скопирован в буфер обмена")
		mergeBtn.Show()
	})

	buttonsContainer := container.NewGridWithColumns(2,
		copyBtn,
		mergeBtn,
	)

	return container.NewBorder(nil, buttonsContainer, nil, nil, entry)
}
