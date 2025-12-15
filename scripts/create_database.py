# create_database.py
from config import DatabaseConfig
import pyodbc

def create_datawarehouse():
    """Crée la base de données du data warehouse si elle n'existe pas"""
    try:
        print("Tentative de création de la base de données NEWW...")
        
        # Connexion au serveur master sans base spécifique
        master_conn_str = (
            'DRIVER={ODBC Driver 18 for SQL Server};'
            'SERVER=localhost;'
            'DATABASE=master;'
            'Trusted_Connection=yes;'
            'Encrypt=no;'
        )
        
        conn = pyodbc.connect(master_conn_str, autocommit=True)  # IMPORTANT: autocommit=True
        cursor = conn.cursor()
        
        # Vérifier si la base existe
        db_name = DatabaseConfig.DW_SERVER['database']
        cursor.execute(f"SELECT name FROM sys.databases WHERE name = '{db_name}'")
        exists = cursor.fetchone()
        
        if not exists:
            print(f"Création de la base de données '{db_name}'...")
            cursor.execute(f"CREATE DATABASE [{db_name}]")
            print(f"✅ Base de données '{db_name}' créée avec succès.")
        else:
            print(f"ℹ️ Base de données '{db_name}' existe déjà.")
        
        cursor.close()
        conn.close()
        
    except Exception as e:
        print(f"❌ Erreur création base de données: {e}")

def create_dw_schema():
    """Crée les tables du data warehouse"""
    try:
        print("\nTentative de connexion à la base NEWW pour créer le schéma...")
        
        # Connexion directe à la base NEWW
        neww_conn_str = (
            'DRIVER={ODBC Driver 18 for SQL Server};'
            'SERVER=localhost;'
            'DATABASE=DataWareHouse;'
            'Trusted_Connection=yes;'
            'Encrypt=no;'
        )
        
        conn = pyodbc.connect(neww_conn_str)
        cursor = conn.cursor()
        print("✅ Connexion à NEWW établie.")
        
        # 1. Table DimDate
        print("Création de DimDate...")
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                          WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'DimDate')
            BEGIN
                CREATE TABLE DimDate (
                    DateKey INT PRIMARY KEY,
                    Date DATE NOT NULL,
                    Year INT NOT NULL,
                    Quarter INT NOT NULL,
                    Month INT NOT NULL,
                    Day INT NOT NULL,
                    MonthName VARCHAR(20),
                    DayOfWeek VARCHAR(20),
                    IsWeekend BIT,
                    UNIQUE(Date)
                );
                PRINT 'Table DimDate créée.';
            END
        """)
        print("✅ DimDate vérifiée/créée.")
        
        # 2. Table DimCustomer
        print("Création de DimCustomer...")
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                          WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'DimCustomer')
            BEGIN
                CREATE TABLE DimCustomer (
                    CustomerKey INT IDENTITY(1,1) PRIMARY KEY,
                    CustomerID VARCHAR(10) NOT NULL,
                    CompanyName VARCHAR(100) NOT NULL,
                    ContactName VARCHAR(100),
                    ContactTitle VARCHAR(100),
                    Address VARCHAR(200),
                    City VARCHAR(50),
                    Region VARCHAR(50),
                    PostalCode VARCHAR(20),
                    Country VARCHAR(50),
                    Phone VARCHAR(30),
                    SourceSystem VARCHAR(20),
                    UNIQUE(CustomerID, SourceSystem)
                );
            END
        """)
        print("✅ DimCustomer vérifiée/créée.")
        
        # 3. Table DimEmployee
        print("Création de DimEmployee...")
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                          WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'DimEmployee')
            BEGIN
                CREATE TABLE DimEmployee (
                    EmployeeKey INT IDENTITY(1,1) PRIMARY KEY,
                    EmployeeID INT NOT NULL,
                    LastName VARCHAR(50) NOT NULL,
                    FirstName VARCHAR(50) NOT NULL,
                    Title VARCHAR(100),
                    TitleOfCourtesy VARCHAR(25),
                    BirthDate DATE,
                    HireDate DATE,
                    Address VARCHAR(200),
                    City VARCHAR(50),
                    Region VARCHAR(50),
                    PostalCode VARCHAR(20),
                    Country VARCHAR(50),
                    HomePhone VARCHAR(30),
                    ReportsTo INT,
                    SourceSystem VARCHAR(20),
                    UNIQUE(EmployeeID, SourceSystem)
                );
            END
        """)
        print("✅ DimEmployee vérifiée/créée.")
        
        # 4. Table FactOrders (sans contraintes de clé étrangère d'abord)
        print("Création de FactOrders...")
        cursor.execute("""
            IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES 
                          WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'FactOrders')
            BEGIN
                CREATE TABLE FactOrders (
                    OrderKey INT IDENTITY(1,1) PRIMARY KEY,
                    OrderID INT NOT NULL,
                    CustomerKey INT NOT NULL,
                    EmployeeKey INT NOT NULL,
                    OrderDateKey INT NOT NULL,
                    RequiredDateKey INT,
                    ShippedDateKey INT,
                    OrderDate DATE,
                    RequiredDate DATE,
                    ShippedDate DATE,
                    ShipVia INT,
                    Freight DECIMAL(10,2),
                    ShipName VARCHAR(100),
                    ShipAddress VARCHAR(200),
                    ShipCity VARCHAR(50),
                    ShipRegion VARCHAR(50),
                    ShipPostalCode VARCHAR(20),
                    ShipCountry VARCHAR(50),
                    IsDelivered BIT DEFAULT 0,
                    DeliveryDelayDays INT,
                    TotalAmount DECIMAL(15,2),
                    SourceSystem VARCHAR(20)
                );
            END
        """)
        print("✅ FactOrders créée (contraintes FK à ajouter plus tard).")
        
        # Ajouter les contraintes de clé étrangère APRÈS la création des tables
        print("\nAjout des contraintes de clés étrangères...")
        
        # Vérifier et ajouter les FK une par une
        try:
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS 
                              WHERE CONSTRAINT_NAME = 'FK_FactOrders_DimCustomer')
                BEGIN
                    ALTER TABLE FactOrders 
                    ADD CONSTRAINT FK_FactOrders_DimCustomer 
                    FOREIGN KEY (CustomerKey) REFERENCES DimCustomer(CustomerKey);
                END
            """)
            print("✅ FK vers DimCustomer ajoutée.")
        except Exception as e:
            print(f"⚠️ FK DimCustomer non ajoutée: {e}")
            
        try:
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS 
                              WHERE CONSTRAINT_NAME = 'FK_FactOrders_DimEmployee')
                BEGIN
                    ALTER TABLE FactOrders 
                    ADD CONSTRAINT FK_FactOrders_DimEmployee 
                    FOREIGN KEY (EmployeeKey) REFERENCES DimEmployee(EmployeeKey);
                END
            """)
            print("✅ FK vers DimEmployee ajoutée.")
        except Exception as e:
            print(f"⚠️ FK DimEmployee non ajoutée: {e}")
        
        # Créer des index pour améliorer les performances
        print("\nCréation des index...")
        try:
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_FactOrders_Dates')
                BEGIN
                    CREATE INDEX IX_FactOrders_Dates ON FactOrders (OrderDateKey, RequiredDateKey, ShippedDateKey);
                END
            """)
            print("✅ Index IX_FactOrders_Dates créé.")
        except Exception as e:
            print(f"⚠️ Index dates non créé: {e}")
            
        try:
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_FactOrders_Customer')
                BEGIN
                    CREATE INDEX IX_FactOrders_Customer ON FactOrders (CustomerKey);
                END
            """)
            print("✅ Index IX_FactOrders_Customer créé.")
        except Exception as e:
            print(f"⚠️ Index customer non créé: {e}")
        
        conn.commit()
        print("\n✅ Schema du data warehouse créé avec succès.")
        
        cursor.close()
        conn.close()
        
    except Exception as e:
        print(f"❌ Erreur création schéma: {e}")

if __name__ == "__main__":
    print("=" * 60)
    print("CRÉATION DU DATA WAREHOUSE")
    print("=" * 60)
    
    # Étape 1: Créer la base de données
    create_datawarehouse()
    
    # Attendre un peu pour laisser SQL Server terminer la création
    import time
    time.sleep(2)
    
    # Étape 2: Créer le schéma
    create_dw_schema()
    
    print("\n" + "=" * 60)
    print("PROCESSUS TERMINÉ")
    print("=" * 60)