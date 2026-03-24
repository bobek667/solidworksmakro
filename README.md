# SolidWorks Macro Center

To jest lekka aplikacja desktopowa w PowerShell z oknem Windows, ktora sluzy jako centrum uruchamiania makr SolidWorks.

## Co potrafi

- wyswietla liste makr z opisami,
- pozwala dodawac, edytowac i usuwac wpisy,
- zapisuje konfiguracje do pliku `macros.json`,
- probuje uruchomic makro bezposrednio przez COM SolidWorks,
- jesli SolidWorks nie jest wlaczony, uruchamia go i czeka na gotowosc,
- pozwala szybko otworzyc plik makra albo jego folder.

## Jak uruchomic

Uruchom w PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\launch-macro-center.ps1
```

Mozesz tez kliknac prawym na plik i wybrac uruchomienie w PowerShell.

## Wazne ustawienie dla kazdego makra

Zeby przycisk `Uruchom w SolidWorks` dzialal poprawnie, wpis makra musi miec uzupelnione:

- `module` - nazwa modulu VBA wewnatrz pliku `.swp`,
- `procedure` - nazwa procedury startowej, np. `main`.

Jesli tych danych jeszcze nie znasz, aplikacja nadal bedzie przydatna jako katalog makr i skrot do plikow. Gdy ustalisz modul i procedure, dopiszesz je w edycji makra i uruchamianie z poziomu programu zacznie dzialac.

## Jak to rozwinac dalej

Kolejny krok moze byc taki:

1. dodamy pola parametrow dla konkretnych makr,
2. dodamy przyjazniejsze kafelki zamiast prostej listy,
3. spakujemy to do `.exe`,
4. podepniemy logi, status wykonania i historie uruchomien.
