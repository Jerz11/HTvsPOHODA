# Aplikace pro srovnání excelů

Tool name: "HT vs Pohoda"
- automatizační tool porovnávající Excel reporty a srovnávající částky stejných dokladů v obou excelech.

Jazyk: Python (pandas + openpyxl pro práci s Excelem)
UI: Streamlit (rychlá tvorba interaktivního webového UI)
Výstup: Přehledná tabulka v UI
Pokročilé možnosti: Možnost exportu do excelu, PDF, interaktivní filtry, vizualizace
Repoty bude možné nahrát pomocí drag and drop nebo pomocí tlačítka "Nahrát soubor"
Uživatel bude spouštět tool pomocí exe souboru

Tool bude srovnávat dva různé excelovské reporty, vyhledá stejná čísla dokladů v obou reportech a porovná částky u těchto dokladů. Doklady, které jsou v obou reportech, budou umístěny v horní části tabulky pod nimi bude jejich součet, ostatní nespárované doklady budou umístěny v dolní části tabulky.

Oba excely, které budou nahrávány, mají pouze jeden sheet.

HT neboli HotelTime reporty mají ve svém názvu "HT" a Pohoda reporty mají ve svém názvu "FV".
Oba reporty mají v prvním řádku headers a číslo dokladu i částky můžou být v různých sloupcích. Číslo dokladu v HT je ve sloupci s názvem "Číslo dokladu". Pohoda má  číslo dokladu ve sloupci s názvem "Číslo".

HT má částku ve sloupci s názvem "Celkem s DPH". 
Pohoda má částku ve sloupci s názvem "Celkem".


V excelech může mít stejný doklad více řádků, takže  bude potřeba nejdříve vyhledat všechny řádky s tímto dokladem a částky posčítat. Jako výstup pak bude pouze jeden řádek s tímto dokladem a jeho částkou.
Jako výstup budeme potřebovat srovnání obou reportů a částek pro všechny doklady, u dokladů, které jsou přítomny v obou reportech, zvýraznit, kde jsou částky jiné. Ostatní doklady, které jsou v jednom nebo druhém reportu, ale nejsou v obou reportech, budou pouze přidány do spodní části tabulky pod řádkem se součtem částek dokladů, které jsou v obou reportech.






