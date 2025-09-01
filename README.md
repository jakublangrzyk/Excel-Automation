# Automatyzacja przyjęcia faktur w Excelu

Excel to nadal jedno z najczęściej stosowanych narzędzi w firmach.  
W tym projekcie pokazuję, jak z pomocą **Power Query** i **tabel przestawnych** można w prosty sposób zautomatyzować proces przyjmowania faktur do systemu.

1. **Ładowanie danych**  
- Power Query pobiera faktury z folderu,  
- narzędzie jest uniwersalne, obsługuje różne formaty (PDF, CSV, XLSX).

2. **Czyszczenie trudnych plików (PDF)**  
- Dane z PDF wymagają dodatkowego czyszczenia, w tym celu korzystam z zaawansowanych formuł w Excelu, aby wyciągnąć potrzebne informacje.

3. **Tabele przestawne i analiza**  
- Czyste dane są ładowane do tabel przestawnych, która zawiera ilości, wartości faktur w przejrzystej formie.

4. ## Pobieranie kursu walut z strony NBP
- Power Query pozwala również na pobieranie danych bezpośrednio ze stron internetowych lub API.  
- W tym projekcie wykorzystałem API NBP, kursy walut są automatycznie pobierane i łączone z fakturami.  
Przy każdym uruchomieniu raportu Power Query pobiera aktualne kursy, dzięki czemu wartości faktur w EUR i USD są od razu przeliczane na PLN.  
![KursEuro](images/KursEuro.png)

5. **Przykłady użycia zaawansowanych formuł**
Formuła wyszukująca określone produkty w danej kolumnie.
```excel
=LET(
  a;[@Column3];
  JEŻELI(
    LUB(
      CZY.LICZBA(ZNAJDŹ("Produkt1";a));
      CZY.LICZBA(ZNAJDŹ("Produkt2";a));
      CZY.LICZBA(ZNAJDŹ("Produkt3";a));
      CZY.LICZBA(ZNAJDŹ("Produkt4";a))
    );
    a;
    ""
  )
)
```

Formuła wyszukująca dane w wielu kolumnach.
```excel
=LET(
  a;[@Column7];
  b;[@Column6];
  c;[@Column5];
  JEŻELI.BŁĄD(
    WYSZUKAJ.PIONOWO(a;SzukanaInformacja!A:A;1;0);
    JEŻELI.BŁĄD(
      WYSZUKAJ.PIONOWO(b;SzukanaInformacja!A:A;1;0);
      JEŻELI.BŁĄD(
        WYSZUKAJ.PIONOWO(c;SzukanaInformacja!A:A;1;0);
        ""
      )
    )
  )
)
```
### Raport końcowy – tabela przestawna
Tabela przedstawia zestawienie liczby sztuk i wartości produktów w podziale na oddziały firmy.  
![CleanData](images/CleanData.png)

## Korzyści 
Taki proces pozwala przyjmować wiele faktur jednocześnie, proces szybszy i skalowalny,
- skrócenie czasu pracy o co najmniej 50% w porównaniu do manualnego wprowadzania danych,  
- automatyczne pobieranie kursów walut brak błędów przy ręcznym wpisywaniu,
- większa przejrzystość i spójność raportów.
