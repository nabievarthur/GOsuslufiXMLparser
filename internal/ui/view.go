package ui

import (
	"fmt"
	"strconv"
	"strings"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/theme"
	"fyne.io/fyne/v2/widget"

	"nabievarthur/GOsuslugiXML/internal/service"
)

func BuildUI(win fyne.Window) fyne.CanvasObject {
	parser := service.NewXMLParser()
	notifier := NewNotifier()

	label1 := widget.NewLabel("Выберите файл XML")
	label1.Alignment = fyne.TextAlignCenter

	label2 := widget.NewLabel("")
	label2.Hide()
	// аккордеон
	accordion := widget.NewAccordion()
	accordion.MultiOpen = true 

	separatorWithPadding := container.NewVBox(
		container.NewWithoutLayout(),
		widget.NewSeparator(),
		container.NewWithoutLayout(),
	)
	separatorWithPadding.Hide()

	var prepareBtn *widget.Button
	//подготовка данных
	prepareBtn = widget.NewButtonWithIcon("Подготовить", theme.ConfirmIcon(), func() {
		notifier.Show("Выполнение...")
		prepareBtn.Hide()

		go func() {
			// Парсим
			res, err := parser.ParseXMLToFile(label1.Text)

			if err != nil {
				fyne.Do(func() {
					notifier.Show("Ошибка: " + err.Error())
				})
				return
			}

			fyne.Do(func() {
				accordion.Items = nil
				// Делим на строки
				lines := strings.Split(strings.TrimSpace(res), "\n")
				totalLines := len(lines)

				// Настройки
				maxLinesPerTab := 500 // Максимальное количество строк на вкладку
				currentLine := 0

				// Создаем вкладки с группами строк
				for currentLine < totalLines {
					endLine := currentLine + maxLinesPerTab
					if endLine > totalLines {
						endLine = totalLines
					}

					//текст для текущей вкладки
					tabLines := lines[currentLine:endLine]
					tabText := strings.Join(tabLines, "\n")

					// Создаем вкладку с содержимым
					tabNumber := len(accordion.Items) + 1
					linesInTab := len(tabLines)

					item := &widget.AccordionItem{
						Title:  fmt.Sprintf("Часть %d (%d строк)", tabNumber, linesInTab),
						Detail: createTabContent(tabText, win, notifier, label1.Text),
					}

					accordion.Append(item)
					currentLine = endLine
				}

				if len(accordion.Items) > 0 {
					accordion.Items[0].Open = true

					accordion.Refresh()

					go func() {
						fyne.Do(func() {
							accordion.Refresh()
							win.Content().Refresh()
						})
					}()
				}

				label2.SetText(fmt.Sprintf("Всего строк: %d, Вкладок: %d", totalLines, len(accordion.Items)))
				label2.Show()
				notifier.Show("Готово! Создано вкладок: " + strconv.Itoa(len(accordion.Items)))
			})
		}()
		separatorWithPadding.Show()
	})
	prepareBtn.Importance = widget.HighImportance
	prepareBtn.Hide()

	//Кнопка загрузки XML файла
	openBtn := widget.NewButtonWithIcon("Выбрать файл выгрузки", theme.FileIcon(), func() {
		fileName := OpenFileDialog()
		if fileName == "" {
			notifier.Show("Файл не выбран")
			return
		}
		label1.SetText(fileName)
		prepareBtn.Show()

		accordion.Items = nil
		accordion.Refresh()
		label2.Hide()
	})

	content := container.NewBorder(nil, notifier.Widget(), nil, nil,
		container.NewVBox(
			label1,
			openBtn,
			prepareBtn,
			separatorWithPadding,
			accordion,
			label2,
		),
	)

	return content
}
