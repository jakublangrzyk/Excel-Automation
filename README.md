# Automatyzacja przyjęcia faktur w Excelu

Excel to nadal jedno z najczęściej stosowanych narzędzi w firmach.  
W tym projekcie pokazuję, jak z pomocą **Power Query** i **tabel przestawnych** można w prosty sposób zautomatyzować proces przyjmowania faktur do systemu.

1. **Ładowanie danych**  
Power Query pobiera faktury z folderu.  
Narzędzie jest uniwersalne, obsługuje różne formaty (PDF, CSV, XLSX).

2. **Czyszczenie trudnych plików (PDF)**  
Dane z PDF wymagają dodatkowego czyszczenia, w tym celu korzystam z zaawansowanych formuł w Excelu, aby wyciągnąć potrzebne informacje.

3. **Tabele przestawne i analiza**  
Czyste dane są ładowane do tabel przestawnych, która zawiera ilości, wartości faktur w przejrzystej formie

4. ## Pobieranie kursu walut z strony NBP
Power Query to świetne narzędzie, które pozwala również na pobieranie danych bezpośrednio ze stron internetowych lub API.
Dzięki temu można w pełni zautomatyzować proces przeliczenia walut dla faktur wystawionych w np. w EUR czy USD.  
W przypadku walut Power Query automatycznie pobiera kursy walut z API NBP i łączy je z fakturami, dodatkowo przy każdym uruchomieniu sam pobiera wartości. 
Dzięki temu wartości faktur w EUR i USD są od razu przeliczane na PLN w raporcie. 

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

Formuła wyszukująca dane w wielu kolumnach
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
W niektórych fakturach dane mogły znajdować się w różnych kolumnach, więc ta formuła rozwiązuje ten problem i pozwala znaleźć potrzebne dane.

### Raport końcowy – tabela przestawna
Tabela przedstawia zestawienie liczby sztuk i wartości produktów w podziale na oddziały firmy.  

## Korzyści 
Taki proces pozwala przyjmować wiele faktur jednocześnie, proces szybszy i skalowalny,
-skrócenie czasu pracy o co najmniej 50% w porównaniu do manualnego wprowadzania danych,  
-automatyczne pobieranie kursów walut brak błędów przy ręcznym wpisywaniu,
-większa przejrzystość i spójność raportów.
