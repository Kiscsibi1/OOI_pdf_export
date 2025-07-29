Adatfeltöltés:
	pdf_export/data <--- clinical goal pdf-ek

	ck_export/data <--- ck excel-ek

Futtatás:
	pdf_extractor.m ---> ha csak a clinical goal-os pdf-ekből kellenek az adatok
	
	ck_export.m ---> megcsinálja a clinical goal exportot és a cyber excel exportot is

Ha hiba van és már megjavítottam, akkor a Command Window-ban (ahol az error meg warningok megjelennek) ---> !git pull