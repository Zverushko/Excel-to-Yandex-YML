package main

import (
	"encoding/xml"
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Константы
const (
	DefaultTemplate = "../yandex_market_template.xlsx"
	DefaultOutput   = "../yandex_market.xml"
)

// Структуры для XML
type YMLCatalog struct {
	XMLName xml.Name `xml:"yml_catalog"`
	Date    string   `xml:"date,attr"`
	Shop    Shop     `xml:"shop"`
}

type Shop struct {
	Name       string     `xml:"name"`
	Company    string     `xml:"company"`
	URL        string     `xml:"url"`
	Currencies Currencies `xml:"currencies"`
	Categories Categories `xml:"categories"`
	Offers     Offers     `xml:"offers"`
}

type Currencies struct {
	Currency []Currency `xml:"currency"`
}

type Currency struct {
	ID   string `xml:"id,attr"`
	Rate string `xml:"rate,attr"`
}

type Categories struct {
	Category []Category `xml:"category"`
}

type Category struct {
	ID       string `xml:"id,attr"`
	ParentID string `xml:"parentId,attr,omitempty"`
	Name     string `xml:",chardata"`
}

type Offers struct {
	Offer []Offer `xml:"offer"`
}

type Offer struct {
	ID          string   `xml:"id,attr"`
	Available   string   `xml:"available,attr"`
	URL         string   `xml:"url"`
	Price       string   `xml:"price"`
	CurrencyID  string   `xml:"currencyId"`
	CategoryID  string   `xml:"categoryId"`
	Picture     []string `xml:"picture,omitempty"`
	Name        string   `xml:"name"`
	Vendor      string   `xml:"vendor,omitempty"`
	Description string   `xml:"description,omitempty"`
	SalesNotes  string   `xml:"sales_notes,omitempty"`
	Params      []Param  `xml:"param,omitempty"`
}

type Param struct {
	Name  string `xml:"name,attr"`
	Unit  string `xml:"unit,attr,omitempty"`
	Value string `xml:",chardata"`
}

// Структуры для хранения данных из Excel
type ShopData struct {
	Name    string
	Company string
	URL     string
}

type ProductData struct {
	ID          string
	Available   bool
	URL         string
	Price       float64
	CurrencyID  string
	CategoryID  string
	Pictures    []string
	Name        string
	Vendor      string
	Description string
	SalesNotes  string
	Params      []ParamData
}

type ParamData struct {
	Name  string
	Unit  string
	Value string
}

// Функция для чтения данных из Excel
func parseExcelToYML(inputFile, outputFile string) error {
	// Открываем Excel файл
	xlsx, err := excelize.OpenFile(inputFile)
	if err != nil {
		return fmt.Errorf("ошибка при открытии Excel файла: %w", err)
	}
	defer xlsx.Close()

	// Читаем данные магазина
	shopData, err := readShopSettings(xlsx)
	if err != nil {
		return fmt.Errorf("ошибка при чтении настроек магазина: %w", err)
	}

	// Читаем валюты
	currencies, err := readCurrencies(xlsx)
	if err != nil {
		return fmt.Errorf("ошибка при чтении валют: %w", err)
	}

	// Читаем категории
	categories, err := readCategories(xlsx)
	if err != nil {
		return fmt.Errorf("ошибка при чтении категорий: %w", err)
	}

	// Читаем товары
	products, err := readProducts(xlsx)
	if err != nil {
		return fmt.Errorf("ошибка при чтении товаров: %w", err)
	}

	// Создаем XML структуру
	ymlCatalog := createYMLCatalog(shopData, currencies, categories, products)

	// Записываем XML в файл
	err = writeXML(ymlCatalog, outputFile)
	if err != nil {
		return fmt.Errorf("ошибка при записи XML: %w", err)
	}

	fmt.Printf("YML-файл успешно создан: %s\n", outputFile)
	return nil
}

// Чтение настроек магазина
func readShopSettings(xlsx *excelize.File) (ShopData, error) {
	var shopData ShopData

	// Проверяем наличие листа
	sheets := xlsx.GetSheetList()
	sheetName := "Настройки магазина"
	sheetExists := false
	for _, sheet := range sheets {
		if sheet == sheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		return shopData, fmt.Errorf("лист '%s' не найден", sheetName)
	}

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return shopData, fmt.Errorf("ошибка при чтении листа '%s': %w", sheetName, err)
	}

	// Пропускаем заголовок
	for i, row := range rows {
		if i == 0 {
			continue // Пропускаем заголовок
		}
		if len(row) >= 2 {
			key := row[0]
			value := row[1]
			switch key {
			case "Название магазина":
				shopData.Name = value
			case "Название компании":
				shopData.Company = value
			case "URL сайта":
				shopData.URL = value
			}
		}
	}

	// Если какие-то поля не заполнены, используем значения по умолчанию
	if shopData.Name == "" {
		shopData.Name = "Мой магазин"
	}
	if shopData.Company == "" {
		shopData.Company = "Моя компания"
	}
	if shopData.URL == "" {
		shopData.URL = "https://example.com"
	}

	return shopData, nil
}

// Чтение валют
func readCurrencies(xlsx *excelize.File) ([]Currency, error) {
	var currencies []Currency

	// Проверяем наличие листа
	sheets := xlsx.GetSheetList()
	sheetName := "Валюты"
	sheetExists := false
	for _, sheet := range sheets {
		if sheet == sheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		return currencies, fmt.Errorf("лист '%s' не найден", sheetName)
	}

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return currencies, fmt.Errorf("ошибка при чтении листа '%s': %w", sheetName, err)
	}

	// Пропускаем заголовок
	for i, row := range rows {
		if i == 0 {
			continue // Пропускаем заголовок
		}
		if len(row) >= 2 {
			id := row[0]
			rate := row[1]
			if id != "" {
				currencies = append(currencies, Currency{ID: id, Rate: rate})
			}
		}
	}

	// Если валюты не указаны, возвращаем ошибку
	if len(currencies) == 0 {
		return currencies, fmt.Errorf("не указаны валюты в листе '%s'", sheetName)
	}

	return currencies, nil
}

// Чтение категорий
func readCategories(xlsx *excelize.File) ([]Category, error) {
	var categories []Category

	// Проверяем наличие листа
	sheets := xlsx.GetSheetList()
	sheetName := "Категории"
	sheetExists := false
	for _, sheet := range sheets {
		if sheet == sheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		return categories, fmt.Errorf("лист '%s' не найден", sheetName)
	}

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return categories, fmt.Errorf("ошибка при чтении листа '%s': %w", sheetName, err)
	}

	// Пропускаем заголовок
	for i, row := range rows {
		if i == 0 {
			continue // Пропускаем заголовок
		}
		if len(row) >= 2 {
			id := row[0]
			name := row[1]
			parentID := ""
			if len(row) >= 3 {
				parentID = row[2]
			}

			if id != "" && name != "" {
				categories = append(categories, Category{
					ID:       id,
					ParentID: parentID,
					Name:     name,
				})
			}
		}
	}

	// Если категории не указаны, возвращаем ошибку
	if len(categories) == 0 {
		return categories, fmt.Errorf("не указаны категории в листе '%s'", sheetName)
	}

	return categories, nil
}

// Чтение товаров
func readProducts(xlsx *excelize.File) ([]ProductData, error) {
	var products []ProductData

	// Проверяем наличие листа
	sheets := xlsx.GetSheetList()
	sheetName := "Товары"
	sheetExists := false
	for _, sheet := range sheets {
		if sheet == sheetName {
			sheetExists = true
			break
		}
	}

	if !sheetExists {
		return products, fmt.Errorf("лист '%s' не найден", sheetName)
	}

	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return products, fmt.Errorf("ошибка при чтении листа '%s': %w", sheetName, err)
	}

	// Определяем индексы колонок
	var colIndexes = make(map[string]int)
	if len(rows) > 0 {
		for i, cell := range rows[0] {
			colIndexes[cell] = i
		}
	} else {
		return products, fmt.Errorf("лист '%s' не содержит данных", sheetName)
	}

	// Пропускаем заголовок
	for i, row := range rows {
		if i == 0 {
			continue // Пропускаем заголовок
		}

		// Проверяем, что строка не пустая
		if len(row) == 0 {
			continue
		}

		// Получаем ID и название товара
		var id, name string

		if idIdx, ok := colIndexes["ID товара"]; ok && len(row) > idIdx {
			id = row[idIdx]
		}

		if nameIdx, ok := colIndexes["Название товара"]; ok && len(row) > nameIdx {
			name = row[nameIdx]
		}

		// Если нет ID или названия, пропускаем товар
		if id == "" || name == "" {
			continue
		}

		product := ProductData{
			ID:   id,
			Name: name,
		}

		// Заполняем остальные поля, если они есть
		if idx, ok := colIndexes["Доступность (true/false)"]; ok && len(row) > idx {
			product.Available = strings.ToLower(row[idx]) == "да" || strings.ToLower(row[idx]) == "true"
		} else {
			product.Available = true // По умолчанию товар доступен
		}

		if idx, ok := colIndexes["URL товара"]; ok && len(row) > idx {
			product.URL = row[idx]
		}

		if idx, ok := colIndexes["Цена"]; ok && len(row) > idx {
			price, err := strconv.ParseFloat(row[idx], 64)
			if err == nil {
				product.Price = price
			}
		}

		if idx, ok := colIndexes["Валюта (ID)"]; ok && len(row) > idx {
			product.CurrencyID = row[idx]
		}

		if idx, ok := colIndexes["ID категории"]; ok && len(row) > idx {
			product.CategoryID = row[idx]
		}

		if idx, ok := colIndexes["URL изображения"]; ok && len(row) > idx {
			pictures := strings.Split(row[idx], ",")
			for _, pic := range pictures {
				pic = strings.TrimSpace(pic)
				if pic != "" {
					product.Pictures = append(product.Pictures, pic)
				}
			}
		}

		if idx, ok := colIndexes["Производитель"]; ok && len(row) > idx {
			product.Vendor = row[idx]
		}

		if idx, ok := colIndexes["Описание"]; ok && len(row) > idx {
			product.Description = row[idx]
		}

		if idx, ok := colIndexes["Примечания"]; ok && len(row) > idx {
			product.SalesNotes = row[idx]
		}

		// Обрабатываем параметры товара
		for key, idx := range colIndexes {
			if strings.HasPrefix(key, "Параметр:") && len(row) > idx {
				paramName := strings.TrimPrefix(key, "Параметр:")
				paramValue := row[idx]
				if paramValue != "" {
					// Проверяем, есть ли единица измерения
					parts := strings.Split(paramName, "(")
					name := strings.TrimSpace(paramName)
					unit := ""
					if len(parts) > 1 {
						name = strings.TrimSpace(parts[0])
						unit = strings.TrimSuffix(strings.TrimSpace(parts[1]), ")")
					}

					product.Params = append(product.Params, ParamData{
						Name:  name,
						Unit:  unit,
						Value: paramValue,
					})
				}
			}
		}

		products = append(products, product)
	}

	// Если товары не указаны, возвращаем ошибку
	if len(products) == 0 {
		return products, fmt.Errorf("не найдено товаров для добавления в YML-файл")
	}

	return products, nil
}

// Создание YML каталога
func createYMLCatalog(shopData ShopData, currencies []Currency, categories []Category, products []ProductData) YMLCatalog {
	// Текущая дата в формате YYYY-MM-DD HH:MM
	now := time.Now().Format("2006-01-02 15:04")

	// Создаем структуру YML каталога
	ymlCatalog := YMLCatalog{
		Date: now,
		Shop: Shop{
			Name:    shopData.Name,
			Company: shopData.Company,
			URL:     shopData.URL,
			Currencies: Currencies{
				Currency: currencies,
			},
			Categories: Categories{
				Category: categories,
			},
			Offers: Offers{},
		},
	}

	// Добавляем товары
	for _, product := range products {
		offer := Offer{
			ID:          product.ID,
			Available:   strconv.FormatBool(product.Available),
			URL:         product.URL,
			Price:       fmt.Sprintf("%.2f", product.Price),
			CurrencyID:  product.CurrencyID,
			CategoryID:  product.CategoryID,
			Picture:     product.Pictures,
			Name:        product.Name,
			Vendor:      product.Vendor,
			Description: product.Description,
			SalesNotes:  product.SalesNotes,
		}

		// Добавляем параметры товара
		for _, param := range product.Params {
			offer.Params = append(offer.Params, Param{
				Name:  param.Name,
				Unit:  param.Unit,
				Value: param.Value,
			})
		}

		ymlCatalog.Shop.Offers.Offer = append(ymlCatalog.Shop.Offers.Offer, offer)
	}

	return ymlCatalog
}

// Запись XML в файл
func writeXML(ymlCatalog YMLCatalog, outputFile string) error {
	// Создаем файл
	file, err := os.Create(outputFile)
	if err != nil {
		return err
	}
	defer file.Close()

	// Записываем XML заголовок
	file.WriteString("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")

	// Создаем XML encoder
	encoder := xml.NewEncoder(file)
	encoder.Indent("", "  ")

	// Кодируем структуру в XML
	if err := encoder.Encode(ymlCatalog); err != nil {
		return err
	}

	return nil
}

func main() {
	// Парсим аргументы командной строки
	inputFile := flag.String("input", DefaultTemplate, "Путь к Excel файлу с данными")
	outputFile := flag.String("output", DefaultOutput, "Путь для сохранения YML файла")
	flag.Parse()

	// Проверяем существование входного файла
	if _, err := os.Stat(*inputFile); os.IsNotExist(err) {
		fmt.Printf("Ошибка: файл %s не существует\n", *inputFile)
		os.Exit(1)
	}

	// Обрабатываем Excel и создаем YML
	err := parseExcelToYML(*inputFile, *outputFile)
	if err != nil {
		fmt.Printf("Ошибка: %s\n", err)
		os.Exit(1)
	}
}
