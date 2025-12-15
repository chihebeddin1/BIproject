import sys
import pandas as pd
import pyodbc
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from datetime import datetime
import numpy as np

# Database connection
def get_db_connection():
    """Establish database connection"""
    try:
        conn_str = (
            'DRIVER={ODBC Driver 18 for SQL Server};'
            'SERVER=localhost;'
            'DATABASE=DataWareHouse;'
            'Trusted_Connection=yes;'
            'Encrypt=no;'
        )
        conn = pyodbc.connect(conn_str)
        return conn
    except Exception as e:
        print(f"Database connection error: {e}")
        return None

class DataWarehouseDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.load_data()
        
    def initUI(self):
        """Initialize the user interface"""
        self.setWindowTitle('📊 Data Warehouse Dashboard')
        self.setGeometry(100, 100, 1400, 900)
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Create header
        self.create_header(main_layout)
        
        # Create main content area
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(15, 15, 15, 15)
        content_layout.setSpacing(15)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setDocumentMode(True)
        content_layout.addWidget(self.tab_widget)
        
        # Create tabs
        self.create_overview_tab()
        self.create_sales_tab()
        self.create_customers_tab()
        self.create_employees_tab()
        self.create_time_analysis_tab()
        self.create_data_explorer_tab()
        
        main_layout.addWidget(content_widget)
        
        # Status bar
        self.statusBar().showMessage('Ready')
        
        # Apply styles
        self.apply_styles()
        
    def create_header(self, parent_layout):
        """Create application header"""
        header_widget = QWidget()
        header_widget.setFixedHeight(80)
        header_widget.setObjectName("headerWidget")  # Add object name for styling
        
        # Apply header-specific style
        header_widget.setStyleSheet("""
            QWidget#headerWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2c3e50, stop:1 #3498db);
                border-bottom: 2px solid #2980b9;
            }
        """)
        
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 10, 20, 10)
        
        # Logo/title section
        title_container = QWidget()
        title_container.setStyleSheet("background: transparent;")
        title_layout = QVBoxLayout(title_container)
        title_layout.setContentsMargins(0, 0, 0, 0)
        
        title_label = QLabel("📊 Data Warehouse Analytics")
        title_label.setStyleSheet("""
            color: white;
            font-size: 20px;
            font-weight: bold;
            padding: 0px;
            background: transparent;
        """)
        
        subtitle_label = QLabel("Real-time Business Intelligence Dashboard")
        subtitle_label.setStyleSheet("""
            color: #ecf0f1;
            font-size: 11px;
            padding: 0px;
            background: transparent;
        """)
        
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        
        header_layout.addWidget(title_container)
        header_layout.addStretch()
        
        # Date and time display
        time_container = QWidget()
        time_container.setStyleSheet("background: transparent;")
        time_layout = QVBoxLayout(time_container)
        time_layout.setContentsMargins(0, 0, 0, 0)
        
        self.date_label = QLabel(datetime.now().strftime("%A, %B %d, %Y"))
        self.date_label.setStyleSheet("""
            color: white;
            font-size: 11px;
            text-align: right;
            padding: 0px;
            background: transparent;
        """)
        
        self.time_label = QLabel(datetime.now().strftime("%I:%M %p"))
        self.time_label.setStyleSheet("""
            color: white;
            font-size: 11px;
            text-align: right;
            padding: 0px;
            background: transparent;
        """)
        
        time_layout.addWidget(self.date_label)
        time_layout.addWidget(self.time_label)
        
        header_layout.addWidget(time_container)
        
        parent_layout.addWidget(header_widget)
        
        # Update time every minute
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_time)
        self.timer.start(60000)  # Update every minute
        
    def update_time(self):
        """Update time display"""
        current_time = datetime.now()
        self.date_label.setText(current_time.strftime("%A, %B %d, %Y"))
        self.time_label.setText(current_time.strftime("%I:%M %p"))
        
    def apply_styles(self):
        """Apply application-wide styles"""
        # Remove the header styling from the global stylesheet
        # and keep only the main window and widget styles
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            
            /* Tab widget styling */
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                margin-top: 2px;
            }
            
            QTabBar::tab {
                background-color: #e9ecef;
                color: #495057;
                padding: 10px 20px;
                margin-right: 3px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: 500;
                font-size: 12px;
                border: 1px solid #dee2e6;
                border-bottom: none;
                min-width: 100px;
            }
            
            QTabBar::tab:selected {
                background-color: white;
                color: #3498db;
                font-weight: 600;
                border-color: #dee2e6;
                border-bottom-color: white;
            }
            
            QTabBar::tab:hover:!selected {
                background-color: #f8f9fa;
            }
            
            /* Button styling */
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
                font-size: 12px;
                min-height: 32px;
            }
            
            QPushButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton:pressed {
                background-color: #1c6ea4;
            }
            
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
            
            /* Frame styling for content frames */
            QFrame[frameShape="4"] {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 6px;
            }
            
            /* Table styling */
            QTableWidget {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                gridline-color: #e9ecef;
                alternate-background-color: #f8f9fa;
                font-size: 11px;
            }
            
            QTableWidget::item {
                padding: 6px;
                border-right: 1px solid #e9ecef;
                border-bottom: 1px solid #e9ecef;
            }
            
            QTableWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            
            QHeaderView::section {
                background-color: #f1f3f4;
                padding: 8px 6px;
                border: 1px solid #dee2e6;
                font-weight: 600;
                color: #2c3e50;
                font-size: 11px;
            }
            
            /* Combo box styling */
            QComboBox {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px;
                min-width: 100px;
                font-size: 12px;
                height: 30px;
            }
            
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            
            QComboBox:hover {
                border-color: #3498db;
            }
            
            /* Date edit styling */
            QDateEdit {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px;
                font-size: 12px;
                height: 30px;
            }
            
            /* Spin box styling */
            QSpinBox {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px;
                font-size: 12px;
                height: 30px;
            }
            
            /* Text edit styling */
            QTextEdit {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', monospace;
                font-size: 11px;
            }
            
            /* Label styling for general labels */
            QLabel {
                color: #2c3e50;
                font-size: 12px;
                background: transparent;
            }
            
            /* Splitter styling */
            QSplitter::handle {
                background-color: #dee2e6;
                width: 3px;
            }
            
            QSplitter::handle:hover {
                background-color: #3498db;
            }
            
            /* Scroll bar styling */
            QScrollBar:vertical {
                border: none;
                background-color: #f8f9fa;
                width: 10px;
                margin: 0px;
            }
            
            QScrollBar::handle:vertical {
                background-color: #bdc3c7;
                border-radius: 4px;
                min-height: 25px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #95a5a6;
            }
            
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            
            /* Status bar */
            QStatusBar {
                background-color: #2c3e50;
                color: white;
                font-size: 10px;
            }
        """)
        
   
    def load_data(self):
        """Load all data from database"""
        try:
            conn = get_db_connection()
            if conn:
                # Load fact orders
                self.orders_df = pd.read_sql("SELECT * FROM FactOrders", conn)
                self.orders_df['OrderDate'] = pd.to_datetime(self.orders_df['OrderDate'])
                
                # Load customers
                self.customers_df = pd.read_sql("SELECT * FROM DimCustomer", conn)
                
                # Load employees
                self.employees_df = pd.read_sql("SELECT * FROM DimEmployee", conn)
                
                # Load date dimension
                self.dates_df = pd.read_sql("SELECT * FROM DimDate", conn)
                self.dates_df['Date'] = pd.to_datetime(self.dates_df['Date'])
                
                conn.close()
                self.statusBar().showMessage('✅ Data loaded successfully')
                
                # Update all tabs with data
                self.update_overview()
                self.update_sales_tab()
                self.update_customers_tab()
                self.update_employees_tab()
                self.update_time_analysis()
                
            else:
                self.statusBar().showMessage('❌ Failed to connect to database')
                
        except Exception as e:
            self.statusBar().showMessage(f'⚠️ Error loading data: {str(e)}')
            
    def create_overview_tab(self):
        """Create Overview tab with key metrics"""
        self.overview_tab = QWidget()
        self.tab_widget.addTab(self.overview_tab, "🏠 Overview")
        
        layout = QVBoxLayout(self.overview_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # Metrics grid
        metrics_widget = QWidget()
        metrics_layout = QGridLayout(metrics_widget)
        metrics_layout.setSpacing(10)
        metrics_layout.setContentsMargins(0, 0, 0, 0)
        
        # Define metrics with icons and colors
        self.metric_widgets = {}
        metrics_config = [
            ('Total Orders', '📦', '#3498db'),
            ('Total Revenue', '💰', '#2ecc71'),
            ('Total Customers', '👥', '#e74c3c'),
            ('Avg Order Value', '📊', '#f39c12'),
            ('Delivered Orders', '✅', '#27ae60'),
            ('Pending Orders', '⏳', '#e67e22'),
            ('Top Customer', '👑', '#9b59b6'),
            ('Top Employee', '⭐', '#1abc9c')
        ]
        
        for i, (title, icon, color) in enumerate(metrics_config):
            metric_card = self.create_metric_card(title, icon, color)
            self.metric_widgets[title] = metric_card
            metrics_layout.addWidget(metric_card, i//4, i%4)
        
        layout.addWidget(metrics_widget)
        
        # Charts section
        charts_widget = QWidget()
        charts_layout = QHBoxLayout(charts_widget)
        charts_layout.setSpacing(15)
        charts_layout.setContentsMargins(0, 0, 0, 0)
        
        # Revenue trend chart
        revenue_container = QWidget()
        revenue_container.setMinimumWidth(400)
        revenue_layout = QVBoxLayout(revenue_container)
        revenue_layout.setContentsMargins(0, 0, 0, 0)
        
        revenue_label = QLabel("📈 Monthly Revenue Trend")
        revenue_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 5px;
        """)
        revenue_layout.addWidget(revenue_label)
        
        self.revenue_canvas = FigureCanvas(Figure(figsize=(6, 3)))
        self.revenue_canvas.setMinimumHeight(250)
        revenue_layout.addWidget(self.revenue_canvas)
        
        charts_layout.addWidget(revenue_container)
        
        # Top customers chart
        customers_container = QWidget()
        customers_container.setMinimumWidth(400)
        customers_layout = QVBoxLayout(customers_container)
        customers_layout.setContentsMargins(0, 0, 0, 0)
        
        customers_label = QLabel("👥 Top 10 Customers")
        customers_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 5px;
        """)
        customers_layout.addWidget(customers_label)
        
        self.customers_canvas = FigureCanvas(Figure(figsize=(6, 3)))
        self.customers_canvas.setMinimumHeight(250)
        customers_layout.addWidget(self.customers_canvas)
        
        charts_layout.addWidget(customers_container)
        
        layout.addWidget(charts_widget)
        layout.addStretch()
        
    def create_metric_card(self, title, icon, color):
        """Create a metric card widget"""
        card = QWidget()
        card.setMinimumHeight(90)
        card.setMaximumHeight(110)
        
        layout = QVBoxLayout(card)
        layout.setContentsMargins(15, 10, 15, 10)
        layout.setSpacing(5)
        
        # Icon and title
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        icon_label = QLabel(icon)
        icon_label.setStyleSheet(f"font-size: 16px; color: {color};")
        header_layout.addWidget(icon_label)
        
        title_label = QLabel(title)
        title_label.setStyleSheet("""
            color: #495057;
            font-size: 12px;
            font-weight: 500;
        """)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        layout.addLayout(header_layout)
        
        # Value
        value_label = QLabel("Loading...")
        value_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 20px;
            font-weight: 600;
            padding: 0px;
        """)
        layout.addWidget(value_label)
        
        layout.addStretch()
        
        # Store reference to value label
        card.value_label = value_label
        
        # Card styling
        card.setStyleSheet(f"""
            QWidget {{
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                border-left: 4px solid {color};
            }}
            
            QWidget:hover {{
                box-shadow: 0 2px 8px rgba(52, 152, 219, 0.1);
            }}
        """)
        
        return card
        
    def update_overview(self):
        """Update overview tab with data"""
        if hasattr(self, 'orders_df'):
            # Calculate metrics
            total_orders = len(self.orders_df)
            total_revenue = self.orders_df['TotalAmount'].sum()
            total_customers = self.customers_df['CustomerID'].nunique() if hasattr(self, 'customers_df') else 0
            avg_order = self.orders_df['TotalAmount'].mean()
            delivered_orders = self.orders_df['IsDelivered'].sum() if 'IsDelivered' in self.orders_df.columns else 0
            pending_orders = total_orders - delivered_orders
            
            # Find top customer
            if hasattr(self, 'orders_df') and hasattr(self, 'customers_df'):
                customer_revenue = self.orders_df.groupby('CustomerKey')['TotalAmount'].sum()
                if not customer_revenue.empty:
                    top_customer_id = customer_revenue.idxmax()
                    top_customer = self.customers_df[self.customers_df['CustomerKey'] == top_customer_id]['CompanyName'].iloc[0]
                else:
                    top_customer = 'N/A'
            else:
                top_customer = 'N/A'
            
            # Find top employee
            if hasattr(self, 'orders_df') and hasattr(self, 'employees_df'):
                employee_revenue = self.orders_df.groupby('EmployeeKey')['TotalAmount'].sum()
                if not employee_revenue.empty:
                    top_employee_id = employee_revenue.idxmax()
                    top_employee = self.employees_df[self.employees_df['EmployeeKey'] == top_employee_id]
                    if not top_employee.empty:
                        top_employee_name = f"{top_employee['FirstName'].iloc[0]} {top_employee['LastName'].iloc[0]}"
                    else:
                        top_employee_name = 'N/A'
                else:
                    top_employee_name = 'N/A'
            else:
                top_employee_name = 'N/A'
            
            # Update metric cards
            metrics = {
                'Total Orders': f"{total_orders:,}",
                'Total Revenue': f"${total_revenue:,.2f}",
                'Total Customers': f"{total_customers:,}",
                'Avg Order Value': f"${avg_order:,.2f}",
                'Delivered Orders': f"{delivered_orders:,}",
                'Pending Orders': f"{pending_orders:,}",
                'Top Customer': top_customer[:15] + ('...' if len(top_customer) > 15 else ''),
                'Top Employee': top_employee_name[:15] + ('...' if len(top_employee_name) > 15 else '')
            }
            
            for title, value in metrics.items():
                if title in self.metric_widgets:
                    self.metric_widgets[title].value_label.setText(value)
            
            # Update charts with enhanced styling
            self.update_revenue_chart()
            self.update_top_customers_chart()
            
    def update_revenue_chart(self):
        """Update monthly revenue chart with enhanced styling"""
        if hasattr(self, 'orders_df') and hasattr(self, 'dates_df'):
            # Merge orders with dates
            merged_df = self.orders_df.merge(
                self.dates_df[['DateKey', 'Year', 'Month', 'MonthName']],
                left_on='OrderDateKey',
                right_on='DateKey',
                how='left'
            )
            
            # Group by month
            monthly_revenue = merged_df.groupby(['Year', 'Month', 'MonthName'])['TotalAmount'].sum().reset_index()
            monthly_revenue = monthly_revenue.sort_values(['Year', 'Month'])
            monthly_revenue['Period'] = monthly_revenue['MonthName'] + ' ' + monthly_revenue['Year'].astype(str)
            
            # Create chart with enhanced styling
            fig = self.revenue_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            # Use a colormap for gradient bars
            colors = plt.cm.Blues(np.linspace(0.4, 0.8, len(monthly_revenue)))
            bars = ax.bar(range(len(monthly_revenue)), monthly_revenue['TotalAmount'], 
                         color=colors, edgecolor='white', linewidth=1.0)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            # Customize labels and titles
            ax.set_xlabel('Month', fontsize=10, fontweight='medium', color='#495057')
            ax.set_ylabel('Revenue ($)', fontsize=10, fontweight='medium', color='#495057')
            ax.set_title('Monthly Revenue Trend', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            
            ax.set_xticks(range(len(monthly_revenue)))
            ax.set_xticklabels(monthly_revenue['Period'], rotation=45, ha='right', fontsize=8)
            
            # Format y-axis
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Add grid
            ax.grid(True, axis='y', alpha=0.2, linestyle='--', linewidth=0.5)
            
            # Remove top and right spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Add value labels on bars
            for bar in bars:
                height = bar.get_height()
                if height > 0:  # Only add label if height > 0
                    ax.text(bar.get_x() + bar.get_width()/2., height + (height * 0.01),
                           f'${height:,.0f}', ha='center', va='bottom', 
                           fontsize=7, fontweight='medium', color='#2c3e50')
            
            fig.tight_layout(pad=1.0)
            self.revenue_canvas.draw()
            
    def update_top_customers_chart(self):
        """Update top customers chart with enhanced styling"""
        if hasattr(self, 'orders_df') and hasattr(self, 'customers_df'):
            # Merge and get top customers
            customer_revenue = self.orders_df.groupby('CustomerKey')['TotalAmount'].sum().reset_index()
            customer_revenue = customer_revenue.merge(
                self.customers_df[['CustomerKey', 'CompanyName']],
                on='CustomerKey',
                how='left'
            )
            
            top_customers = customer_revenue.nlargest(10, 'TotalAmount')
            
            # Create chart with enhanced styling
            fig = self.customers_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            # Use gradient colors
            colors = plt.cm.Greens(np.linspace(0.4, 0.8, len(top_customers)))
            bars = ax.barh(range(len(top_customers)), top_customers['TotalAmount'], 
                          color=colors, edgecolor='white', linewidth=1.0)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            # Customize labels and titles
            ax.set_xlabel('Total Revenue ($)', fontsize=10, fontweight='medium', color='#495057')
            ax.set_title('Top 10 Customers by Revenue', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            
            ax.set_yticks(range(len(top_customers)))
            ax.set_yticklabels([name[:20] + ('...' if len(name) > 20 else '') 
                               for name in top_customers['CompanyName']], fontsize=8)
            
            # Format x-axis
            ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Add grid
            ax.grid(True, axis='x', alpha=0.2, linestyle='--', linewidth=0.5)
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Add value labels
            for bar in bars:
                width = bar.get_width()
                if width > 0:  # Only add label if width > 0
                    ax.text(width + (width * 0.01), bar.get_y() + bar.get_height()/2.,
                           f'${width:,.0f}', ha='left', va='center', 
                           fontsize=7, fontweight='medium', color='#2c3e50')
            
            ax.invert_yaxis()  # Highest value on top
            fig.tight_layout(pad=1.0)
            self.customers_canvas.draw()
            
    def create_sales_tab(self):
        """Create Sales Analytics tab"""
        self.sales_tab = QWidget()
        self.tab_widget.addTab(self.sales_tab, "📈 Sales Analytics")
        
        layout = QVBoxLayout(self.sales_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Filters frame
        filters_frame = QFrame()
        filters_frame.setFrameShape(QFrame.StyledPanel)
        filters_layout = QHBoxLayout(filters_frame)
        filters_layout.setContentsMargins(15, 10, 15, 10)
        
        # Date range filter
        filters_layout.addWidget(QLabel("Date Range:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-3))
        self.start_date_edit.setCalendarPopup(True)
        filters_layout.addWidget(self.start_date_edit)
        
        filters_layout.addWidget(QLabel("to"))
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        filters_layout.addWidget(self.end_date_edit)
        
        # Country filter
        filters_layout.addWidget(QLabel("Country:"))
        self.country_combo = QComboBox()
        self.country_combo.addItem("All Countries")
        self.country_combo.setMinimumWidth(120)
        filters_layout.addWidget(self.country_combo)
        
        # Apply button
        self.apply_filter_btn = QPushButton("Apply Filters")
        self.apply_filter_btn.clicked.connect(self.update_sales_tab)
        filters_layout.addWidget(self.apply_filter_btn)
        
        filters_layout.addStretch()
        layout.addWidget(filters_frame)
        
        # Charts frame
        charts_frame = QFrame()
        charts_frame.setFrameShape(QFrame.StyledPanel)
        charts_layout = QHBoxLayout(charts_frame)
        charts_layout.setContentsMargins(10, 10, 10, 10)
        
        # Sales by country chart
        self.country_chart_canvas = FigureCanvas(Figure(figsize=(6, 4)))
        self.country_chart_canvas.setMinimumWidth(400)
        self.country_chart_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.country_chart_canvas)
        
        # Daily sales chart
        self.daily_chart_canvas = FigureCanvas(Figure(figsize=(6, 4)))
        self.daily_chart_canvas.setMinimumWidth(400)
        self.daily_chart_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.daily_chart_canvas)
        
        layout.addWidget(charts_frame)
        
        # Sales data table
        table_frame = QFrame()
        table_frame.setFrameShape(QFrame.StyledPanel)
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(5, 5, 5, 5)
        
        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(7)
        self.sales_table.setHorizontalHeaderLabels([
            'Order ID', 'Date', 'Customer', 'Country', 'Amount', 'Freight', 'Status'
        ])
        table_layout.addWidget(self.sales_table)
        
        layout.addWidget(table_frame)
        
    def update_sales_tab(self):
        """Update sales tab with filtered data"""
        if hasattr(self, 'orders_df'):
            # Apply filters
            start_date = self.start_date_edit.date().toString('yyyy-MM-dd')
            end_date = self.end_date_edit.date().toString('yyyy-MM-dd')
            
            filtered_df = self.orders_df[
                (self.orders_df['OrderDate'] >= start_date) & 
                (self.orders_df['OrderDate'] <= end_date)
            ]
            
            # Filter by country if selected
            country = self.country_combo.currentText()
            if country != "All Countries" and hasattr(self, 'customers_df'):
                customer_countries = self.customers_df[['CustomerKey', 'Country']]
                filtered_df = filtered_df.merge(customer_countries, on='CustomerKey', how='left')
                filtered_df = filtered_df[filtered_df['Country'] == country]
            
            # Update country combo box
            if hasattr(self, 'customers_df') and self.country_combo.count() == 1:
                countries = self.customers_df['Country'].dropna().unique()
                self.country_combo.addItems(countries)
            
            # Update charts
            self.update_country_chart(filtered_df)
            self.update_daily_sales_chart(filtered_df)
            
            # Update table
            self.update_sales_table(filtered_df)
            
    def update_country_chart(self, df):
        """Update sales by country chart"""
        if hasattr(self, 'customers_df'):
            # Merge with customer data to get countries
            df_with_country = df.merge(
                self.customers_df[['CustomerKey', 'Country']],
                on='CustomerKey',
                how='left'
            )
            
            country_sales = df_with_country.groupby('Country')['TotalAmount'].sum().reset_index()
            country_sales = country_sales.sort_values('TotalAmount', ascending=False).head(10)
            
            fig = self.country_chart_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            colors = plt.cm.Set3(np.linspace(0, 1, len(country_sales)))
            bars = ax.bar(country_sales['Country'], country_sales['TotalAmount'], 
                         color=colors, edgecolor='white', linewidth=1.0)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            ax.set_xlabel('Country', fontsize=10, color='#495057')
            ax.set_ylabel('Revenue ($)', fontsize=10, color='#495057')
            ax.set_title('Top 10 Countries by Revenue', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            ax.set_xticklabels(country_sales['Country'], rotation=45, ha='right', fontsize=8)
            
            # Format y-axis
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Add grid
            ax.grid(True, axis='y', alpha=0.2, linestyle='--', linewidth=0.5)
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Add value labels
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height + (height * 0.01),
                           f'${height:,.0f}', ha='center', va='bottom', 
                           fontsize=7, fontweight='medium', color='#2c3e50')
            
            fig.tight_layout(pad=1.0)
            self.country_chart_canvas.draw()
            
    def update_daily_sales_chart(self, df):
        """Update daily sales chart"""
        if not df.empty:
            daily_sales = df.groupby(df['OrderDate'].dt.date)['TotalAmount'].sum().reset_index()
            daily_sales.columns = ['Date', 'Revenue']
            daily_sales = daily_sales.sort_values('Date')
            
            fig = self.daily_chart_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            ax.plot(daily_sales['Date'], daily_sales['Revenue'], 
                   color='#3498db', linewidth=2.0, marker='o', markersize=4)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            ax.set_xlabel('Date', fontsize=10, color='#495057')
            ax.set_ylabel('Revenue ($)', fontsize=10, color='#495057')
            ax.set_title('Daily Sales Trend', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            ax.grid(True, alpha=0.2, linestyle='--')
            
            # Format axes
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%m-%d'))
            ax.xaxis.set_major_locator(plt.matplotlib.dates.AutoDateLocator())
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            fig.tight_layout(pad=1.0)
            self.daily_chart_canvas.draw()
            
    def update_sales_table(self, df):
        """Update sales data table"""
        self.sales_table.setRowCount(min(100, len(df)))  # Show max 100 rows
        
        for i, row in df.head(100).iterrows():
            # Get customer name
            customer_name = "Unknown"
            if hasattr(self, 'customers_df'):
                customer = self.customers_df[self.customers_df['CustomerKey'] == row['CustomerKey']]
                if not customer.empty:
                    customer_name = customer['CompanyName'].iloc[0]
            
            # Get country
            country = "Unknown"
            if hasattr(self, 'customers_df'):
                customer = self.customers_df[self.customers_df['CustomerKey'] == row['CustomerKey']]
                if not customer.empty:
                    country = customer['Country'].iloc[0]
            
            # Populate table
            items = [
                str(row['OrderID']),
                str(row['OrderDate'])[:10],
                customer_name,
                country,
                f"${row['TotalAmount']:,.2f}",
                f"${row['Freight']:,.2f}" if pd.notna(row['Freight']) else "$0.00",
                "Delivered" if row.get('IsDelivered', 0) == 1 else "Pending"
            ]
            
            for j, item in enumerate(items):
                self.sales_table.setItem(i, j, QTableWidgetItem(item))
        
        self.sales_table.resizeColumnsToContents()
        
    def create_customers_tab(self):
        """Create Customer Insights tab"""
        self.customers_tab = QWidget()
        self.tab_widget.addTab(self.customers_tab, "👥 Customers")
        
        layout = QVBoxLayout(self.customers_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Filters
        filters_frame = QFrame()
        filters_frame.setFrameShape(QFrame.StyledPanel)
        filters_layout = QHBoxLayout(filters_frame)
        filters_layout.setContentsMargins(15, 10, 15, 10)
        
        filters_layout.addWidget(QLabel("Segment:"))
        self.segment_combo = QComboBox()
        self.segment_combo.addItems(["All", "High Value", "Medium", "Low"])
        filters_layout.addWidget(self.segment_combo)
        
        filters_layout.addWidget(QLabel("Min Orders:"))
        self.min_orders_spin = QSpinBox()
        self.min_orders_spin.setMinimum(0)
        self.min_orders_spin.setMaximum(1000)
        self.min_orders_spin.setValue(1)
        filters_layout.addWidget(self.min_orders_spin)
        
        filters_layout.addWidget(QLabel("Country:"))
        self.customer_country_combo = QComboBox()
        self.customer_country_combo.addItem("All")
        self.customer_country_combo.setMinimumWidth(120)
        filters_layout.addWidget(self.customer_country_combo)
        
        self.apply_customer_filter_btn = QPushButton("Apply")
        self.apply_customer_filter_btn.clicked.connect(self.update_customers_tab)
        filters_layout.addWidget(self.apply_customer_filter_btn)
        
        filters_layout.addStretch()
        layout.addWidget(filters_frame)
        
        # Charts frame
        charts_frame = QFrame()
        charts_frame.setFrameShape(QFrame.StyledPanel)
        charts_layout = QHBoxLayout(charts_frame)
        charts_layout.setContentsMargins(10, 10, 10, 10)
        
        # Customer segmentation chart
        self.segmentation_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        self.segmentation_canvas.setMinimumWidth(350)
        self.segmentation_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.segmentation_canvas)
        
        # Country distribution chart
        self.customer_country_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        self.customer_country_canvas.setMinimumWidth(350)
        self.customer_country_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.customer_country_canvas)
        
        layout.addWidget(charts_frame)
        
        # Customer table
        table_frame = QFrame()
        table_frame.setFrameShape(QFrame.StyledPanel)
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(5, 5, 5, 5)
        
        self.customers_table = QTableWidget()
        self.customers_table.setColumnCount(6)
        self.customers_table.setHorizontalHeaderLabels([
            'Company', 'Country', 'City', 'Contact', 'Orders', 'Total Spent'
        ])
        table_layout.addWidget(self.customers_table)
        
        layout.addWidget(table_frame)
        
    def update_customers_tab(self):
        """Update customers tab"""
        if hasattr(self, 'customers_df') and hasattr(self, 'orders_df'):
            # Calculate customer metrics
            customer_orders = self.orders_df.groupby('CustomerKey').agg({
                'OrderID': 'count',
                'TotalAmount': 'sum'
            }).reset_index()
            customer_orders.columns = ['CustomerKey', 'OrderCount', 'TotalSpent']
            
            # Merge with customer data
            customers_full = self.customers_df.merge(
                customer_orders,
                on='CustomerKey',
                how='left'
            )
            customers_full['OrderCount'] = customers_full['OrderCount'].fillna(0)
            customers_full['TotalSpent'] = customers_full['TotalSpent'].fillna(0)
            
            # Apply filters
            filtered_customers = customers_full.copy()
            
            # Country filter
            country = self.customer_country_combo.currentText()
            if country != "All":
                filtered_customers = filtered_customers[filtered_customers['Country'] == country]
            
            # Min orders filter
            min_orders = self.min_orders_spin.value()
            filtered_customers = filtered_customers[filtered_customers['OrderCount'] >= min_orders]
            
            # Segment filter
            segment = self.segment_combo.currentText()
            if segment != "All":
                # Create segments based on spending
                percentiles = filtered_customers['TotalSpent'].quantile([0.33, 0.67])
                if segment == "High Value":
                    filtered_customers = filtered_customers[filtered_customers['TotalSpent'] > percentiles.iloc[1]]
                elif segment == "Medium":
                    filtered_customers = filtered_customers[
                        (filtered_customers['TotalSpent'] > percentiles.iloc[0]) & 
                        (filtered_customers['TotalSpent'] <= percentiles.iloc[1])
                    ]
                else:  # Low
                    filtered_customers = filtered_customers[filtered_customers['TotalSpent'] <= percentiles.iloc[0]]
            
            # Update country combo box
            if self.customer_country_combo.count() == 1:
                countries = self.customers_df['Country'].dropna().unique()
                self.customer_country_combo.addItems(countries)
            
            # Update charts
            self.update_segmentation_chart(filtered_customers)
            self.update_customer_country_chart(filtered_customers)
            
            # Update table
            self.update_customers_table(filtered_customers)
            
    def update_segmentation_chart(self, df):
        """Update customer segmentation chart"""
        if not df.empty:
            # Create segments
            df['Segment'] = pd.qcut(df['TotalSpent'], q=3, labels=['Low', 'Medium', 'High'])
            segment_counts = df['Segment'].value_counts()
            
            fig = self.segmentation_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            colors = ['#FF9800', '#4CAF50', '#2196F3']
            wedges, texts, autotexts = ax.pie(
                segment_counts.values, 
                labels=segment_counts.index,
                colors=colors,
                autopct='%1.1f%%',
                startangle=90
            )
            
            ax.set_title('Customer Segmentation by Spending', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            ax.axis('equal')
            
            fig.patch.set_facecolor('white')
            fig.tight_layout(pad=1.0)
            self.segmentation_canvas.draw()
            
    def update_customer_country_chart(self, df):
        """Update customer country distribution chart"""
        if not df.empty and 'Country' in df.columns:
            country_counts = df['Country'].value_counts().head(10)
            
            fig = self.customer_country_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            colors = plt.cm.Purples(np.linspace(0.4, 0.8, len(country_counts)))
            bars = ax.bar(range(len(country_counts)), country_counts.values, 
                         color=colors, edgecolor='white', linewidth=1.0)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            ax.set_xlabel('Country', fontsize=10, color='#495057')
            ax.set_ylabel('Number of Customers', fontsize=10, color='#495057')
            ax.set_title('Top 10 Countries by Customer Count', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            ax.set_xticks(range(len(country_counts)))
            ax.set_xticklabels(country_counts.index, rotation=45, ha='right', fontsize=8)
            
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            ax.grid(True, axis='y', alpha=0.2, linestyle='--', linewidth=0.5)
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Add value labels
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height + (height * 0.01),
                           f'{int(height)}', ha='center', va='bottom', 
                           fontsize=7, fontweight='medium', color='#2c3e50')
            
            fig.tight_layout(pad=1.0)
            self.customer_country_canvas.draw()
            
    def update_customers_table(self, df):
        """Update customers table"""
        self.customers_table.setRowCount(min(100, len(df)))
        
        for i, row in df.head(100).iterrows():
            items = [
                str(row.get('CompanyName', '')),
                str(row.get('Country', '')),
                str(row.get('City', '')),
                str(row.get('ContactName', '')),
                str(int(row.get('OrderCount', 0))),
                f"${row.get('TotalSpent', 0):,.2f}"
            ]
            
            for j, item in enumerate(items):
                self.customers_table.setItem(i, j, QTableWidgetItem(item))
        
        self.customers_table.resizeColumnsToContents()
        
    def create_employees_tab(self):
        """Create fully dynamic Employee Performance tab"""
        self.employees_tab = QWidget()
        self.tab_widget.addTab(self.employees_tab, "👨‍💼 Employees")
        
        # Main layout with stretch factors
        layout = QVBoxLayout(self.employees_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Performance metrics - Dynamic grid
        metrics_container = QWidget()
        metrics_container.setMinimumHeight(100)
        metrics_container.setMaximumHeight(150)
        metrics_layout = QGridLayout(metrics_container)
        metrics_layout.setContentsMargins(10, 10, 10, 10)
        metrics_layout.setSpacing(10)
        
        self.employee_metric_labels = {}
        metric_config = [
            ('Top Employee', '👑', '#3498db'),
            ('Highest Revenue', '💰', '#2ecc71'), 
            ('Most Orders', '📦', '#e74c3c'),
            ('Best Avg Order', '⭐', '#f39c12')
        ]
        
        for i, (title, icon, color) in enumerate(metric_config):
            metric_card = self.create_dynamic_metric_card(title, icon, color)
            self.employee_metric_labels[title] = metric_card.value_label
            metrics_layout.addWidget(metric_card, i//2, i%2)
        
        layout.addWidget(metrics_container)
        
        # Charts section with splitter for dynamic resizing
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)
        splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #dee2e6;
                border-radius: 3px;
            }
            QSplitter::handle:hover {
                background-color: #3498db;
            }
        """)
        
        # Left chart - Employee performance
        left_chart_container = QWidget()
        left_layout = QVBoxLayout(left_chart_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        perf_label = QLabel("📊 Employee Performance")
        perf_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 5px;
        """)
        left_layout.addWidget(perf_label)
        
        self.employee_perf_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        left_layout.addWidget(self.employee_perf_canvas)
        
        splitter.addWidget(left_chart_container)
        
        # Right chart - Performance by title
        right_chart_container = QWidget()
        right_layout = QVBoxLayout(right_chart_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        title_label = QLabel("👔 Performance by Title")
        title_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 5px;
        """)
        right_layout.addWidget(title_label)
        
        self.title_perf_canvas = FigureCanvas(Figure(figsize=(5, 4)))
        right_layout.addWidget(self.title_perf_canvas)
        
        splitter.addWidget(right_chart_container)
        
        # Set initial sizes (can be adjusted by user)
        splitter.setSizes([400, 400])
        layout.addWidget(splitter, 3)  # Stretch factor 3 for charts
        
        # Employee table with dynamic resizing
        table_container = QWidget()
        table_layout = QVBoxLayout(table_container)
        table_layout.setContentsMargins(0, 0, 0, 0)
        
        table_label = QLabel("📋 Employee Data")
        table_label.setStyleSheet("""
            color: #2c3e50;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 5px;
        """)
        table_layout.addWidget(table_label)
        
        self.employees_table = QTableWidget()
        self.employees_table.setColumnCount(6)
        self.employees_table.setHorizontalHeaderLabels([
            'Name', 'Title', 'Country', 'Orders', 'Revenue', 'Avg Order'
        ])
        
        # Make table headers resizable
        header = self.employees_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)
        
        # Enable sorting
        self.employees_table.setSortingEnabled(True)
        
        table_layout.addWidget(self.employees_table)
        layout.addWidget(table_container, 2)  # Stretch factor 2 for table
        
    def create_dynamic_metric_card(self, title, icon, color):
        """Create a dynamic metric card that adjusts to size"""
        card = QWidget()
        card.setMinimumHeight(80)
        
        layout = QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(5)
        
        # Top row: icon and title
        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(8)
        
        icon_label = QLabel(icon)
        icon_label.setStyleSheet(f"""
            font-size: 16px;
            color: {color};
            min-width: 24px;
        """)
        top_row.addWidget(icon_label)
        
        title_label = QLabel(title)
        title_label.setStyleSheet("""
            color: #495057;
            font-size: 11px;
            font-weight: 500;
        """)
        title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        top_row.addWidget(title_label)
        
        layout.addLayout(top_row)
        
        # Value label with dynamic font size
        value_label = QLabel("Loading...")
        value_label.setStyleSheet(f"""
            color: #2c3e50;
            font-weight: 600;
            padding: 2px 0px;
        """)
        value_label.setWordWrap(True)
        value_label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        value_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        # Store font size adjustment
        value_label.original_font_size = 14
        
        layout.addWidget(value_label)
        
        # Store reference
        card.value_label = value_label
        
        # Card styling
        card.setStyleSheet(f"""
            QWidget {{
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                border-left: 4px solid {color};
            }}
            QWidget:hover {{
                box-shadow: 0 2px 8px rgba(52, 152, 219, 0.1);
            }}
        """)
        
        return card
    
    def update_employees_tab(self):
        """Update employees tab with dynamic content"""
        if hasattr(self, 'employees_df') and hasattr(self, 'orders_df'):
            # Calculate employee performance
            employee_perf = self.orders_df.groupby('EmployeeKey').agg({
                'OrderID': 'count',
                'TotalAmount': 'sum'
            }).reset_index()
            employee_perf.columns = ['EmployeeKey', 'OrderCount', 'TotalRevenue']
            employee_perf['AvgOrder'] = employee_perf['TotalRevenue'] / employee_perf['OrderCount']
            
            # Merge with employee data
            employees_full = self.employees_df.merge(
                employee_perf,
                on='EmployeeKey',
                how='left'
            ).fillna(0)
            
            # Update metrics with dynamic text
            if not employees_full.empty:
                # Top employee by revenue
                top_employee = employees_full.loc[employees_full['TotalRevenue'].idxmax()]
                top_employee_name = f"{top_employee['FirstName']} {top_employee['LastName']}"
                self.employee_metric_labels['Top Employee'].setText(top_employee_name)
                
                # Highest revenue
                highest_rev = employees_full['TotalRevenue'].max()
                self.employee_metric_labels['Highest Revenue'].setText(f"${highest_rev:,.0f}")
                
                # Most orders
                most_orders = employees_full.loc[employees_full['OrderCount'].idxmax()]
                most_orders_name = f"{most_orders['FirstName']} {most_orders['LastName']}"
                order_count = int(most_orders['OrderCount'])
                self.employee_metric_labels['Most Orders'].setText(
                    f"{most_orders_name}\n({order_count} orders)"
                )
                
                # Best average order
                best_avg = employees_full.loc[employees_full['AvgOrder'].idxmax()]
                self.employee_metric_labels['Best Avg Order'].setText(f"${best_avg['AvgOrder']:,.0f}")
            
            # Adjust font sizes based on content length
            self.adjust_metric_font_sizes()
            
            # Update charts with dynamic data
            self.update_employee_performance_chart(employees_full)
            self.update_title_performance_chart(employees_full)
            
            # Update table
            self.update_employees_table(employees_full)
            
            # Update charts when tab is resized
            self.employees_tab.resizeEvent = lambda event: self.on_employees_tab_resize(event, employees_full)
    
    def adjust_metric_font_sizes(self):
        """Dynamically adjust font sizes based on content length"""
        for title, label in self.employee_metric_labels.items():
            text = label.text()
            length = len(text)
            
            # Adjust font size based on text length
            if length > 30:
                font_size = 10
            elif length > 20:
                font_size = 11
            elif length > 15:
                font_size = 12
            else:
                font_size = label.original_font_size if hasattr(label, 'original_font_size') else 14
            
            current_style = label.styleSheet()
            # Update font size in stylesheet
            new_style = current_style.replace(f"font-size: {label.original_font_size}px", f"font-size: {font_size}px")
            if "font-size:" not in current_style:
                new_style += f"font-size: {font_size}px;"
            
            label.setStyleSheet(new_style)
            label.setWordWrap(length > 15)  # Enable word wrap for longer texts
    
    def on_employees_tab_resize(self, event, df):
        """Handle tab resize events"""
        QWidget.resizeEvent(self.employees_tab, event)
        
        # Redraw charts with updated dimensions
        self.update_employee_performance_chart(df)
        self.update_title_performance_chart(df)
        
        # Adjust table column widths
        self.employees_table.resizeColumnsToContents()
        
        # Adjust metric font sizes
        self.adjust_metric_font_sizes()
    
    def update_employee_performance_chart(self, df):
        """Update employee performance chart with dynamic sizing"""
        if not df.empty:
            fig = self.employee_perf_canvas.figure
            fig.clf()
            
            # Get current canvas size
            canvas_width = self.employee_perf_canvas.width() / 100  # Convert to inches
            canvas_height = self.employee_perf_canvas.height() / 100
            
            # Adjust figure size based on canvas
            fig.set_size_inches(max(4, canvas_width * 0.9), max(3, canvas_height * 0.9))
            
            ax = fig.add_subplot(111)
            
            # Filter out zeros for better visualization
            plot_df = df[(df['OrderCount'] > 0) & (df['TotalRevenue'] > 0)]
            
            if not plot_df.empty:
                # Dynamic point sizing based on data range
                max_avg = plot_df['AvgOrder'].max() if plot_df['AvgOrder'].max() > 0 else 100
                point_sizes = 50 + (plot_df['AvgOrder'] / max_avg * 150)
                
                scatter = ax.scatter(
                    plot_df['OrderCount'], 
                    plot_df['TotalRevenue'],
                    s=point_sizes,
                    c=plot_df['AvgOrder'],
                    cmap='viridis',
                    alpha=0.7,
                    edgecolors='white',
                    linewidth=0.5
                )
                
                # Dynamic font sizing based on chart size
                title_font_size = max(10, min(14, int(canvas_height * 2)))
                label_font_size = max(8, min(12, int(canvas_height * 1.5)))
                tick_font_size = max(7, min(10, int(canvas_height * 1.2)))
                
                ax.set_xlabel('Number of Orders', fontsize=label_font_size, color='#495057')
                ax.set_ylabel('Total Revenue ($)', fontsize=label_font_size, color='#495057')
                ax.set_title('Employee Performance Analysis', fontsize=title_font_size, 
                           fontweight='bold', color='#2c3e50', pad=10)
                
                ax.grid(True, alpha=0.2, linestyle='--')
                ax.tick_params(axis='both', colors='#6c757d', labelsize=tick_font_size)
                
                # Format y-axis
                ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
                
                # Remove spines
                for spine in ['top', 'right']:
                    ax.spines[spine].set_visible(False)
                for spine in ['left', 'bottom']:
                    ax.spines[spine].set_color('#dee2e6')
                
                # Add colorbar with dynamic size
                cbar = plt.colorbar(scatter, ax=ax, pad=0.01)
                cbar.set_label('Average Order Value ($)', fontsize=label_font_size - 1)
                
                # Dynamic annotation - only show if there's enough space
                if canvas_width > 5:  # Only annotate if chart is wide enough
                    top_5 = plot_df.nlargest(min(5, len(plot_df)), 'TotalRevenue')
                    for _, row in top_5.iterrows():
                        ax.annotate(
                            f"{row['FirstName'][0]}.{row['LastName']}",
                            (row['OrderCount'], row['TotalRevenue']),
                            fontsize=max(6, int(tick_font_size * 0.8)),
                            ha='center',
                            bbox=dict(boxstyle="round,pad=0.3", facecolor="white", 
                                    edgecolor="#dee2e6", alpha=0.9)
                        )
            else:
                # No data message
                ax.text(0.5, 0.5, 'No performance data available', 
                       ha='center', va='center', transform=ax.transAxes,
                       fontsize=12, color='#6c757d')
                ax.set_xticks([])
                ax.set_yticks([])
                for spine in ax.spines.values():
                    spine.set_visible(False)
            
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            fig.tight_layout(pad=1.0)
            self.employee_perf_canvas.draw()
    
    def update_title_performance_chart(self, df):
        """Update performance by job title chart with dynamic sizing"""
        if not df.empty and 'Title' in df.columns:
            title_perf = df.groupby('Title').agg({
                'TotalRevenue': 'mean',
                'OrderCount': 'mean'
            }).reset_index()
            
            title_perf = title_perf[title_perf['TotalRevenue'] > 0]
            
            if not title_perf.empty:
                fig = self.title_perf_canvas.figure
                fig.clf()
                
                # Get current canvas size
                canvas_width = self.title_perf_canvas.width() / 100
                canvas_height = self.title_perf_canvas.height() / 100
                
                # Adjust figure size based on canvas
                fig.set_size_inches(max(4, canvas_width * 0.9), max(3, canvas_height * 0.9))
                
                ax = fig.add_subplot(111)
                
                # Sort by revenue for better display
                title_perf = title_perf.sort_values('TotalRevenue', ascending=True)
                
                colors = plt.cm.Greens(np.linspace(0.4, 0.8, len(title_perf)))
                bars = ax.barh(range(len(title_perf)), title_perf['TotalRevenue'], 
                             color=colors, edgecolor='white', linewidth=1.0)
                
                # Dynamic font sizing
                title_font_size = max(10, min(14, int(canvas_height * 2)))
                label_font_size = max(8, min(12, int(canvas_height * 1.5)))
                tick_font_size = max(7, min(10, int(canvas_height * 1.2)))
                
                ax.set_yticks(range(len(title_perf)))
                
                # Truncate long titles if needed
                y_labels = []
                for title in title_perf['Title']:
                    if len(title) > 20 and canvas_width < 6:
                        y_labels.append(title[:18] + '...')
                    else:
                        y_labels.append(title)
                
                ax.set_yticklabels(y_labels, fontsize=tick_font_size)
                ax.set_xlabel('Average Revenue ($)', fontsize=label_font_size, color='#495057')
                ax.set_title('Performance by Job Title', fontsize=title_font_size, 
                           fontweight='bold', color='#2c3e50', pad=10)
                
                ax.invert_yaxis()  # Highest value on top
                ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
                ax.tick_params(axis='both', colors='#6c757d', labelsize=tick_font_size)
                ax.grid(True, axis='x', alpha=0.2, linestyle='--', linewidth=0.5)
                
                # Remove spines
                for spine in ['top', 'right']:
                    ax.spines[spine].set_visible(False)
                for spine in ['left', 'bottom']:
                    ax.spines[spine].set_color('#dee2e6')
                
                # Dynamic value labels - only show if there's enough space
                if canvas_width > 4:
                    for bar in bars:
                        width = bar.get_width()
                        if width > 0:
                            ax.text(width + (width * 0.01), bar.get_y() + bar.get_height()/2.,
                                   f'${width:,.0f}', ha='left', va='center', 
                                   fontsize=max(6, int(tick_font_size * 0.8)), 
                                   fontweight='medium', color='#2c3e50')
            else:
                # No data message
                fig = self.title_perf_canvas.figure
                fig.clf()
                ax = fig.add_subplot(111)
                ax.text(0.5, 0.5, 'No title performance data', 
                       ha='center', va='center', transform=ax.transAxes,
                       fontsize=12, color='#6c757d')
                ax.set_xticks([])
                ax.set_yticks([])
                for spine in ax.spines.values():
                    spine.set_visible(False)
            
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            fig.tight_layout(pad=1.0)
            self.title_perf_canvas.draw()
    
    def update_employees_table(self, df):
        """Update employees table with dynamic column widths"""
        df = df.sort_values('TotalRevenue', ascending=False)
        self.employees_table.setRowCount(min(50, len(df)))
        
        for i, row in df.head(50).iterrows():
            items = [
                f"{row.get('FirstName', '')} {row.get('LastName', '')}",
                str(row.get('Title', '')),
                str(row.get('Country', '')),
                str(int(row.get('OrderCount', 0))),
                f"${row.get('TotalRevenue', 0):,.2f}",
                f"${row.get('AvgOrder', 0):,.2f}"
            ]
            
            for j, item in enumerate(items):
                self.employees_table.setItem(i, j, QTableWidgetItem(item))
        
        # Adjust column widths based on content
        self.employees_table.resizeColumnsToContents()
        
        # Set minimum column widths
        for col in range(self.employees_table.columnCount()):
            current_width = self.employees_table.columnWidth(col)
            min_width = 80 if col < 3 else 100
            if current_width < min_width:
                self.employees_table.setColumnWidth(col, min_width)
                
    def create_time_analysis_tab(self):
        """Create Time Analysis tab"""
        self.time_tab = QWidget()
        self.tab_widget.addTab(self.time_tab, "📅 Time Analysis")
        
        layout = QVBoxLayout(self.time_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Period selector with Apply and Refresh buttons
        period_frame = QFrame()
        period_frame.setFrameShape(QFrame.StyledPanel)
        period_layout = QHBoxLayout(period_frame)
        period_layout.setContentsMargins(15, 10, 15, 10)
        
        period_layout.addWidget(QLabel("Analysis Period:"))
        self.period_combo = QComboBox()
        self.period_combo.addItems(["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"])
        # Remove auto-update when selection changes
        # self.period_combo.currentTextChanged.connect(self.update_time_analysis)
        period_layout.addWidget(self.period_combo)
        
        period_layout.addWidget(QLabel("Year:"))
        self.year_combo = QComboBox()
        self.year_combo.addItem("All Years")
        # Remove auto-update when selection changes
        # self.year_combo.currentTextChanged.connect(self.update_time_analysis)
        period_layout.addWidget(self.year_combo)
        
        # Add Apply button
        self.apply_time_filter_btn = QPushButton("Apply Filters")
        self.apply_time_filter_btn.clicked.connect(self.update_time_analysis)
        period_layout.addWidget(self.apply_time_filter_btn)
        
        # Add Refresh button
        self.refresh_time_btn = QPushButton("Refresh")
        self.refresh_time_btn.clicked.connect(self.refresh_time_analysis)
        period_layout.addWidget(self.refresh_time_btn)
        
        period_layout.addStretch()
        layout.addWidget(period_frame)
        
        # Charts frame
        charts_frame = QFrame()
        charts_frame.setFrameShape(QFrame.StyledPanel)
        charts_layout = QVBoxLayout(charts_frame)
        charts_layout.setContentsMargins(10, 10, 10, 10)
        
        # Time series chart
        self.time_series_canvas = FigureCanvas(Figure(figsize=(10, 4)))
        self.time_series_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.time_series_canvas)
        
        # Day of week analysis
        self.dow_canvas = FigureCanvas(Figure(figsize=(10, 4)))
        self.dow_canvas.setMinimumHeight(300)
        charts_layout.addWidget(self.dow_canvas)
        
        layout.addWidget(charts_frame)
        
    def refresh_time_analysis(self):
        """Refresh time analysis to default settings"""
        # Reset combos to default values
        self.period_combo.setCurrentText("Monthly")
        self.year_combo.setCurrentText("All Years")
        
        # Update charts
        self.update_time_analysis()
    def update_time_analysis(self):
        """Update time analysis tab"""
        if hasattr(self, 'orders_df') and hasattr(self, 'dates_df'):
            # Merge data
            df = self.orders_df.merge(
                self.dates_df,
                left_on='OrderDateKey',
                right_on='DateKey',
                how='left'
            )
            
            # Filter by year if selected
            year = self.year_combo.currentText()
            if year != "All Years":
                df = df[df['Year'] == int(year)]
            
            # Update year combo box
            if self.year_combo.count() == 1:
                years = sorted(self.dates_df['Year'].dropna().unique())
                self.year_combo.addItems([str(y) for y in years])
            
            # Update charts based on period
            period = self.period_combo.currentText()
            self.update_time_series_chart(df, period)
            self.update_dow_chart(df)
            
    def update_time_series_chart(self, df, period):
        """Update time series chart based on period"""
        if not df.empty:
            fig = self.time_series_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            if period == "Daily":
                time_data = df.groupby('Date')['TotalAmount'].sum().reset_index()
                time_data = time_data.sort_values('Date')
                x = time_data['Date']
                title = 'Daily Sales'
                
            elif period == "Weekly":
                df['Week'] = df['Date'].dt.isocalendar().week
                time_data = df.groupby(['Year', 'Week'])['TotalAmount'].sum().reset_index()
                time_data = time_data.sort_values(['Year', 'Week'])
                time_data['Period'] = time_data['Year'].astype(str) + '-W' + time_data['Week'].astype(str)
                x = time_data['Period']
                title = 'Weekly Sales'
                
            elif period == "Monthly":
                time_data = df.groupby(['Year', 'Month', 'MonthName'])['TotalAmount'].sum().reset_index()
                time_data = time_data.sort_values(['Year', 'Month'])
                time_data['Period'] = time_data['MonthName'] + ' ' + time_data['Year'].astype(str)
                x = time_data['Period']
                title = 'Monthly Sales'
                
            elif period == "Quarterly":
                time_data = df.groupby(['Year', 'Quarter'])['TotalAmount'].sum().reset_index()
                time_data = time_data.sort_values(['Year', 'Quarter'])
                time_data['Period'] = 'Q' + time_data['Quarter'].astype(str) + ' ' + time_data['Year'].astype(str)
                x = time_data['Period']
                title = 'Quarterly Sales'
                
            else:  # Yearly
                time_data = df.groupby('Year')['TotalAmount'].sum().reset_index()
                time_data = time_data.sort_values('Year')
                x = time_data['Year'].astype(str)
                title = 'Yearly Sales'
            
            # Create enhanced line chart
            ax.plot(x, time_data['TotalAmount'], color='#3498db', linewidth=2.0, 
                   marker='o', markersize=4)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            ax.set_xlabel('Period', fontsize=10, color='#495057')
            ax.set_ylabel('Revenue ($)', fontsize=10, color='#495057')
            ax.set_title(f'{title} Trend', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            ax.grid(True, alpha=0.2, linestyle='--')
            
            # Format y-axis
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Rotate x labels for better readability
            plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
            
            fig.tight_layout(pad=1.0)
            self.time_series_canvas.draw()
            
    def update_dow_chart(self, df):
        """Update day of week analysis chart"""
        if not df.empty and 'DayOfWeek' in df.columns:
            dow_mapping = {
                'Monday': 0, 'Tuesday': 1, 'Wednesday': 2, 
                'Thursday': 3, 'Friday': 4, 'Saturday': 5, 'Sunday': 6
            }
            df['DOW_Num'] = df['DayOfWeek'].map(dow_mapping)
            
            dow_analysis = df.groupby(['DayOfWeek', 'DOW_Num']).agg({
                'TotalAmount': ['mean', 'sum', 'count']
            }).reset_index()
            dow_analysis = dow_analysis.sort_values('DOW_Num')
            
            fig = self.dow_canvas.figure
            fig.clf()
            ax = fig.add_subplot(111)
            
            colors = plt.cm.Set3(np.linspace(0, 1, len(dow_analysis)))
            bars = ax.bar(dow_analysis['DayOfWeek'], dow_analysis[('TotalAmount', 'mean')], 
                         color=colors, edgecolor='white', linewidth=1.0)
            
            # Enhance chart appearance
            ax.set_facecolor('#f8f9fa')
            fig.patch.set_facecolor('white')
            
            ax.set_xlabel('Day of Week', fontsize=10, color='#495057')
            ax.set_ylabel('Average Revenue ($)', fontsize=10, color='#495057')
            ax.set_title('Average Revenue by Day of Week', fontsize=12, fontweight='bold', color='#2c3e50', pad=10)
            
            ax.grid(True, axis='y', alpha=0.2, linestyle='--', linewidth=0.5)
            ax.tick_params(axis='both', colors='#6c757d', labelsize=8)
            
            # Format y-axis
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            
            # Remove spines
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#dee2e6')
            ax.spines['bottom'].set_color('#dee2e6')
            
            # Add value labels
            for bar in bars:
                height = bar.get_height()
                if height > 0:
                    ax.text(bar.get_x() + bar.get_width()/2., height + (height * 0.01),
                           f'${height:,.0f}', ha='center', va='bottom', 
                           fontsize=7, fontweight='medium', color='#2c3e50')
            
            fig.tight_layout(pad=1.0)
            self.dow_canvas.draw()
            
    def create_data_explorer_tab(self):
        """Create Data Explorer tab"""
        self.explorer_tab = QWidget()
        self.tab_widget.addTab(self.explorer_tab, "🔍 Data Explorer")
        
        layout = QVBoxLayout(self.explorer_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Query input
        query_frame = QFrame()
        query_frame.setFrameShape(QFrame.StyledPanel)
        query_layout = QVBoxLayout(query_frame)
        query_layout.setContentsMargins(15, 10, 15, 10)
        
        query_layout.addWidget(QLabel("SQL Query:"))
        self.query_text = QTextEdit()
        self.query_text.setPlaceholderText("Enter your SQL query here...")
        self.query_text.setMaximumHeight(100)
        query_layout.addWidget(self.query_text)
        
        # Sample queries
        sample_queries = QComboBox()
        sample_queries.addItems([
            "Select a sample query...",
            "SELECT TOP 100 * FROM FactOrders",
            "SELECT * FROM DimCustomer",
            "SELECT * FROM DimEmployee",
            "SELECT * FROM DimDate",
            "SELECT CompanyName, Country, COUNT(*) as OrderCount FROM FactOrders fo JOIN DimCustomer dc ON fo.CustomerKey = dc.CustomerKey GROUP BY CompanyName, Country ORDER BY OrderCount DESC"
        ])
        sample_queries.currentTextChanged.connect(self.load_sample_query)
        query_layout.addWidget(sample_queries)
        
        # Buttons
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        button_layout.setContentsMargins(0, 5, 0, 0)
        
        self.execute_btn = QPushButton("Execute Query")
        self.execute_btn.clicked.connect(self.execute_query)
        button_layout.addWidget(self.execute_btn)
        
        self.export_csv_btn = QPushButton("Export to CSV")
        self.export_csv_btn.clicked.connect(self.export_to_csv)
        button_layout.addWidget(self.export_csv_btn)
        
        self.export_excel_btn = QPushButton("Export to Excel")
        self.export_excel_btn.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_excel_btn)
        
        button_layout.addStretch()
        query_layout.addWidget(button_frame)
        
        layout.addWidget(query_frame)
        
        # Results table
        table_frame = QFrame()
        table_frame.setFrameShape(QFrame.StyledPanel)
        table_layout = QVBoxLayout(table_frame)
        table_layout.setContentsMargins(5, 5, 5, 5)
        
        self.results_table = QTableWidget()
        table_layout.addWidget(self.results_table)
        
        layout.addWidget(table_frame)
        
        # Status label
        self.results_label = QLabel("Ready to execute queries")
        self.results_label.setStyleSheet("color: #6c757d; font-size: 11px; padding: 5px;")
        layout.addWidget(self.results_label)
        
    def load_sample_query(self, query):
        """Load sample query into editor"""
        if query and "Select a sample query..." not in query:
            self.query_text.setText(query)
            
    def execute_query(self):
        """Execute SQL query and display results"""
        query = self.query_text.toPlainText().strip()
        if not query:
            QMessageBox.warning(self, "Warning", "Please enter a SQL query")
            return
            
        try:
            conn = get_db_connection()
            if conn:
                df = pd.read_sql(query, conn)
                conn.close()
                
                self.display_query_results(df)
                self.results_label.setText(f"Query executed successfully. {len(df)} rows returned.")
            else:
                QMessageBox.critical(self, "Error", "Could not connect to database")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Query execution failed: {str(e)}")
            
    def display_query_results(self, df):
        """Display query results in table"""
        self.results_table.setRowCount(len(df))
        self.results_table.setColumnCount(len(df.columns))
        self.results_table.setHorizontalHeaderLabels(df.columns)
        
        for i, row in df.iterrows():
            for j, col in enumerate(df.columns):
                value = str(row[col])
                if pd.isna(row[col]):
                    value = "NULL"
                self.results_table.setItem(i, j, QTableWidgetItem(value))
        
        self.results_table.resizeColumnsToContents()
        
    def export_to_csv(self):
        """Export current results to CSV"""
        if self.results_table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No data to export")
            return
            
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save CSV File", "", "CSV Files (*.csv)")
        
        if filename:
            try:
                data = []
                for i in range(self.results_table.rowCount()):
                    row = []
                    for j in range(self.results_table.columnCount()):
                        item = self.results_table.item(i, j)
                        row.append(item.text() if item else "")
                    data.append(row)
                
                headers = []
                for j in range(self.results_table.columnCount()):
                    headers.append(self.results_table.horizontalHeaderItem(j).text())
                
                df = pd.DataFrame(data, columns=headers)
                df.to_csv(filename, index=False)
                QMessageBox.information(self, "Success", f"Data exported to {filename}")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")
                
    def export_to_excel(self):
        """Export current results to Excel"""
        if self.results_table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No data to export")
            return
            
        filename, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "", "Excel Files (*.xlsx)")
        
        if filename:
            try:
                data = []
                for i in range(self.results_table.rowCount()):
                    row = []
                    for j in range(self.results_table.columnCount()):
                        item = self.results_table.item(i, j)
                        row.append(item.text() if item else "")
                    data.append(row)
                
                headers = []
                for j in range(self.results_table.columnCount()):
                    headers.append(self.results_table.horizontalHeaderItem(j).text())
                
                df = pd.DataFrame(data, columns=headers)
                df.to_excel(filename, index=False)
                QMessageBox.information(self, "Success", f"Data exported to {filename}")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application font
    font = QFont("Segoe UI", 9)
    app.setFont(font)
    
    dashboard = DataWarehouseDashboard()
    dashboard.show()
    
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

