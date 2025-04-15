package CustomTheme

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/theme"
	"image/color"
)

type Loader struct {
	Theme   fyne.Theme
	Variant fyne.ThemeVariant
}

func (m Loader) Font(style fyne.TextStyle) fyne.Resource {
	return theme.DefaultTheme().Font(style)
}

func (m Loader) Icon(name fyne.ThemeIconName) fyne.Resource {
	//if name == theme.IconNameHome {
	//	return fyne.NewStaticResource("myHome", homeBytes)
	//}

	return theme.DefaultTheme().Icon(name)
}

func (m Loader) Size(name fyne.ThemeSizeName) float32 {
	return theme.DefaultTheme().Size(name)
}

func (m Loader) Color(name fyne.ThemeColorName, variant fyne.ThemeVariant) color.Color {
	return theme.DefaultTheme().Color(name, m.Variant)
}

var _ fyne.Theme = (*Loader)(nil)
