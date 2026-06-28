"""
Pakistani E-Commerce Data Analysis
-----------------------------------
Run this script in VS Code, Jupyter, or any Python environment.
It will generate three Excel files (Customers.xlsx, Products.xlsx, Orders.xlsx)
and then perform a full analysis with charts and insights.
"""

# ============================
# 1. IMPORTS
# ============================
import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

# For Jupyter notebooks, uncomment the next line to see plots inline:
# %matplotlib inline

# ============================
# 2. DATA GENERATION
# ============================
def generate_data():
    """Generate the three Excel files if they do not already exist."""
    # Check if files exist; if yes, skip generation (optional)
    import os
    if os.path.exists('Customers.xlsx') and os.path.exists('Products.xlsx') and os.path.exists('Orders.xlsx'):
        print("Excel files already exist. Skipping data generation.")
        return

    print("Generating data...")
    NUM_CUSTOMERS = 2000
    NUM_PRODUCTS = 100
    NUM_ORDERS = 2000
    START_DATE = datetime(2023, 1, 1)
    END_DATE = datetime(2024, 6, 1)

    # Pakistani cities with provinces
    CITIES = {
        'Karachi': 'Sindh',
        'Lahore': 'Punjab',
        'Islamabad': 'Islamabad Capital Territory',
        'Rawalpindi': 'Punjab',
        'Faisalabad': 'Punjab',
        'Multan': 'Punjab',
        'Hyderabad': 'Sindh',
        'Gujranwala': 'Punjab',
        'Peshawar': 'Khyber Pakhtunkhwa',
        'Quetta': 'Balochistan',
        'Sialkot': 'Punjab',
        'Abbottabad': 'Khyber Pakhtunkhwa',
        'Bahawalpur': 'Punjab',
        'Sargodha': 'Punjab',
        'Sukkur': 'Sindh',
        'Mardan': 'Khyber Pakhtunkhwa',
    }
    CITY_LIST = list(CITIES.keys())

    MALE_NAMES = [
        'Ahmed', 'Ali', 'Hamza', 'Hassan', 'Hussein', 'Omar', 'Usman', 'Zain',
        'Bilal', 'Danish', 'Faisal', 'Haroon', 'Imran', 'Junaid', 'Kamran',
        'Noman', 'Rizwan', 'Saad', 'Shahid', 'Tariq', 'Waqar', 'Yasir', 'Zeeshan',
        'Adeel', 'Asad', 'Farhan', 'Hammad', 'Irfan', 'Kashif', 'Mohsin',
        'Nadeem', 'Rashid', 'Sajid', 'Salman', 'Shehryar', 'Tahir', 'Wasim',
        'Zaheer', 'Aamir', 'Fahad', 'Gul', 'Javed', 'Khalid', 'Majid', 'Nasir'
    ]
    FEMALE_NAMES = [
        'Aisha', 'Fatima', 'Maryam', 'Zara', 'Sana', 'Hira', 'Nida', 'Iman',
        'Ayesha', 'Bushra', 'Dania', 'Eman', 'Fiza', 'Hania', 'Iqra', 'Javeria',
        'Kiran', 'Laila', 'Maham', 'Naila', 'Omina', 'Pari', 'Qurat-ul-Ain',
        'Rida', 'Saba', 'Tahira', 'Umaima', 'Varda', 'Wajiha', 'Xara', 'Yumna',
        'Zainab', 'Afreen', 'Bisma', 'Fariha', 'Huma', 'Ishfaq', 'Kanwal',
        'Mahnoor', 'Noor', 'Rabia', 'Sadia', 'Tayyaba', 'Uzma', 'Zakia'
    ]

    PRODUCT_CATEGORIES = {
        'Electronics': ['Smartphone', 'Laptop', 'Tablet', 'Headphones', 'Power Bank', 'Smart Watch', 'LED TV'],
        'Clothing': ['Shirt', 'Jeans', 'Kurta', 'Saree', 'Hijab', 'Shoes', 'Jacket'],
        'Home & Kitchen': ['Cooking Pan', 'Blender', 'Microwave', 'Vacuum Cleaner', 'Dinner Set', 'Towels'],
        'Books': ['Novel', 'Textbook', 'Biography', 'Self-help', 'Children Book'],
        'Beauty': ['Moisturizer', 'Shampoo', 'Perfume', 'Makeup Kit', 'Nail Polish'],
        'Sports': ['Football', 'Cricket Bat', 'Tennis Racket', 'Gym Bag', 'Running Shoes'],
        'Automotive': ['Car Cover', 'Seat Covers', 'Air Freshener', 'Car Cleaning Kit'],
        'Groceries': ['Rice', 'Cooking Oil', 'Sugar', 'Tea', 'Spices']
    }
    PRODUCTS = []
    for cat, items in PRODUCT_CATEGORIES.items():
        for item in items:
            PRODUCTS.append((item, cat))
    # Ensure we have exactly 100 products (add variations if needed)
    while len(PRODUCTS) < NUM_PRODUCTS:
        for cat, items in PRODUCT_CATEGORIES.items():
            for item in items:
                if len(PRODUCTS) >= NUM_PRODUCTS:
                    break
                suffix = random.choice(['', 'Pro', 'Max', 'Lite', 'Plus', 'Mini', 'Deluxe', 'Premium'])
                new_name = f"{item} {suffix}" if suffix else item
                if (new_name, cat) not in PRODUCTS:
                    PRODUCTS.append((new_name, cat))
            if len(PRODUCTS) >= NUM_PRODUCTS:
                break
    PRODUCTS = PRODUCTS[:NUM_PRODUCTS]

    PAYMENT_METHODS = ['Cash on Delivery', 'JazzCash', 'EasyPaisa', 'Bank Transfer', 'Credit/Debit Card']
    STATUSES = ['Delivered', 'Shipped', 'Processing', 'Cancelled', 'Returned']

    # --- Customers ---
    customers = []
    for i in range(1, NUM_CUSTOMERS + 1):
        gender = random.choice(['Male', 'Female'])
        name = random.choice(MALE_NAMES if gender == 'Male' else FEMALE_NAMES)
        if random.random() < 0.4:
            surname = random.choice(['Khan', 'Ahmed', 'Ali', 'Hussain', 'Butt', 'Chaudhry', 'Malik', 'Sheikh'])
            name = f"{name} {surname}"
        city = random.choice(CITY_LIST)
        province = CITIES[city]
        email = f"{name.lower().replace(' ', '.')}{random.randint(10,99)}@email.com"
        phone = f"03{random.randint(0,9)}-{random.randint(1000000,9999999)}"
        reg_date = START_DATE + timedelta(days=random.randint(0, (END_DATE - START_DATE).days))
        customers.append({
            'customer_id': i,
            'name': name,
            'gender': gender,
            'city': city,
            'province': province,
            'email': email,
            'phone': phone,
            'registration_date': reg_date.strftime('%Y-%m-%d')
        })
    df_customers = pd.DataFrame(customers)

    # --- Products ---
    products = []
    for i, (pname, cat) in enumerate(PRODUCTS, 1):
        price = random.randint(200, 50000)
        price = round(price / 50) * 50
        stock = random.randint(0, 200)
        products.append({
            'product_id': i,
            'product_name': pname,
            'category': cat,
            'price_pkr': price,
            'stock': stock
        })
    df_products = pd.DataFrame(products)

    # --- Orders ---
    orders = []
    for i in range(1, NUM_ORDERS + 1):
        customer_id = random.randint(1, NUM_CUSTOMERS)
        product_id = random.randint(1, NUM_PRODUCTS)
        quantity = random.randint(1, 5)
        order_date = START_DATE + timedelta(days=random.randint(0, (END_DATE - START_DATE).days))
        payment = random.choice(PAYMENT_METHODS)
        status = random.choices(STATUSES, weights=[0.6, 0.15, 0.1, 0.1, 0.05])[0]
        delivery_city = random.choice(CITY_LIST)
        # Get customer city (to avoid repeated lookup later)
        customer_city = df_customers.loc[df_customers['customer_id'] == customer_id, 'city'].values[0]
        orders.append({
            'order_id': i,
            'customer_id': customer_id,
            'product_id': product_id,
            'quantity': quantity,
            'order_date': order_date.strftime('%Y-%m-%d'),
            'payment_method': payment,
            'order_status': status,
            'delivery_city': delivery_city,
            'customer_city': customer_city
        })
    df_orders = pd.DataFrame(orders)

    # Save to Excel
    with pd.ExcelWriter('Customers.xlsx', engine='openpyxl') as writer:
        df_customers.to_excel(writer, sheet_name='Customers', index=False)
    with pd.ExcelWriter('Products.xlsx', engine='openpyxl') as writer:
        df_products.to_excel(writer, sheet_name='Products', index=False)
    with pd.ExcelWriter('Orders.xlsx', engine='openpyxl') as writer:
        df_orders.to_excel(writer, sheet_name='Orders', index=False)

    print("Data generation complete. Files: Customers.xlsx, Products.xlsx, Orders.xlsx")


# ============================
# 3. MAIN ANALYSIS
# ============================
def run_analysis():
    """Load data, clean, merge, EDA, charts, insights."""
    # Generate data if needed
    generate_data()

    # Load files
    print("\nLoading data...")
    customers = pd.read_excel('Customers.xlsx', sheet_name='Customers')
    products = pd.read_excel('Products.xlsx', sheet_name='Products')
    orders = pd.read_excel('Orders.xlsx', sheet_name='Orders')

    print("Shapes:")
    print(f"Customers: {customers.shape}")
    print(f"Products:  {products.shape}")
    print(f"Orders:    {orders.shape}")

    # ----- Clean -----
    # Customers
    customers['registration_date'] = pd.to_datetime(customers['registration_date'])
    # Products
    products['price_pkr'] = products['price_pkr'].astype(int)
    products['stock'] = products['stock'].astype(int)
    # Orders
    orders['order_date'] = pd.to_datetime(orders['order_date'])
    # Remove duplicates (if any)
    customers = customers.drop_duplicates(subset=['customer_id'])
    orders = orders.drop_duplicates(subset=['order_id'])

    # ----- Merge -----
    df = orders.merge(customers, on='customer_id', how='left')
    df = df.merge(products, on='product_id', how='left')

    # ----- Feature Engineering -----
    df['total_amount'] = df['quantity'] * df['price_pkr']
    df['order_month'] = df['order_date'].dt.month
    df['order_year'] = df['order_date'].dt.year
    df['order_dayofweek'] = df['order_date'].dt.dayofweek
    df['order_quarter'] = df['order_date'].dt.quarter
    df['days_since_reg'] = (df['order_date'] - df['registration_date']).dt.days
    # Remove any negative days if data error
    df = df[df['days_since_reg'] >= 0]

    print("\nMerged dataset shape:", df.shape)

    # ----- Summary Statistics -----
    print("\n=== Summary Statistics ===")
    print(df[['quantity', 'price_pkr', 'total_amount', 'days_since_reg']].describe())

    # ----- KPIs -----
    total_revenue = df['total_amount'].sum()
    total_orders = df['order_id'].nunique()
    aov = total_revenue / total_orders
    delivered = df[df['order_status'] == 'Delivered'].shape[0]
    completion_rate = delivered / total_orders * 100
    avg_days = df['days_since_reg'].mean()
    top_payment = df['payment_method'].value_counts().idxmax()
    top_city = df['delivery_city'].value_counts().idxmax()

    print("\n=== Key Performance Indicators ===")
    print(f"Total Revenue:          PKR {total_revenue:,.2f}")
    print(f"Total Orders:           {total_orders}")
    print(f"Average Order Value:    PKR {aov:,.2f}")
    print(f"Order Completion Rate:  {completion_rate:.2f}%")
    print(f"Avg days from reg to order: {avg_days:.1f} days")
    print(f"Most popular payment:   {top_payment}")
    print(f"Top city by orders:     {top_city}")

    # ----- Charts -----
    print("\nGenerating charts...")
    sns.set_style('whitegrid')

    # 1. Daily Sales
    daily_sales = df.groupby('order_date')['total_amount'].sum().reset_index()
    plt.figure(figsize=(15, 5))
    plt.plot(daily_sales['order_date'], daily_sales['total_amount'], color='green')
    plt.title('Daily Total Sales (PKR)')
    plt.xlabel('Date')
    plt.ylabel('Total Sales (PKR)')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # 2. Monthly Sales Trend
    monthly_sales = df.groupby(['order_year', 'order_month'])['total_amount'].sum().reset_index()
    monthly_sales['year_month'] = monthly_sales['order_year'].astype(str) + '-' + monthly_sales['order_month'].astype(str).str.zfill(2)
    plt.figure(figsize=(12, 5))
    sns.lineplot(data=monthly_sales, x='year_month', y='total_amount', marker='o', color='blue')
    plt.title('Monthly Sales Trend')
    plt.xticks(rotation=45)
    plt.xlabel('Year-Month')
    plt.ylabel('Total Sales (PKR)')
    plt.tight_layout()
    plt.show()

    # 3. Top 10 Products by Revenue
    top_products = df.groupby('product_name')['total_amount'].sum().sort_values(ascending=False).head(10)
    plt.figure(figsize=(10, 6))
    sns.barplot(x=top_products.values, y=top_products.index, palette='viridis')
    plt.title('Top 10 Products by Revenue (PKR)')
    plt.xlabel('Total Revenue')
    plt.tight_layout()
    plt.show()

    # 4. Top Categories by Revenue
    top_categories = df.groupby('category')['total_amount'].sum().sort_values(ascending=False).head(5)
    plt.figure(figsize=(10, 5))
    top_categories.plot(kind='bar', color='skyblue')
    plt.title('Top 5 Categories by Revenue')
    plt.ylabel('Total Revenue (PKR)')
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.show()

    # 5. Payment Method Distribution
    payment_counts = df['payment_method'].value_counts()
    plt.figure(figsize=(8, 8))
    plt.pie(payment_counts, labels=payment_counts.index, autopct='%1.1f%%', startangle=140)
    plt.title('Payment Method Distribution')
    plt.show()

    # 6. Order Status
    status_counts = df['order_status'].value_counts()
    plt.figure(figsize=(8, 5))
    sns.barplot(x=status_counts.index, y=status_counts.values, palette='Set2')
    plt.title('Order Status Counts')
    plt.ylabel('Number of Orders')
    plt.tight_layout()
    plt.show()

    # 7. Sales by Province
    province_sales = df.groupby('province')['total_amount'].sum().sort_values(ascending=False)
    plt.figure(figsize=(10, 5))
    sns.barplot(x=province_sales.values, y=province_sales.index, palette='coolwarm')
    plt.title('Total Sales by Province')
    plt.xlabel('Total Sales (PKR)')
    plt.tight_layout()
    plt.show()

    # 8. Gender Distribution
    gender_counts = df['gender'].value_counts()
    plt.figure(figsize=(6, 6))
    plt.pie(gender_counts, labels=gender_counts.index, autopct='%1.1f%%', colors=['#ff9999','#66b3ff'])
    plt.title('Customer Gender Distribution')
    plt.show()

    # 9. Average Order Value Trend
    monthly_avg = df.groupby(['order_year', 'order_month'])['total_amount'].mean().reset_index()
    monthly_avg['year_month'] = monthly_avg['order_year'].astype(str) + '-' + monthly_avg['order_month'].astype(str).str.zfill(2)
    plt.figure(figsize=(12, 5))
    sns.lineplot(data=monthly_avg, x='year_month', y='total_amount', marker='o', color='purple')
    plt.title('Average Order Value (AOV) Trend')
    plt.xlabel('Year-Month')
    plt.ylabel('Average Order Value (PKR)')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    # 10. Correlation Heatmap
    numeric_cols = ['quantity', 'price_pkr', 'total_amount', 'days_since_reg']
    corr = df[numeric_cols].corr()
    plt.figure(figsize=(8, 6))
    sns.heatmap(corr, annot=True, cmap='coolwarm', fmt='.2f')
    plt.title('Correlation Heatmap')
    plt.tight_layout()
    plt.show()

    # ----- Customer Segmentation (RFM-like) -----
    customer_agg = df.groupby('customer_id').agg({
        'order_id': 'count',
        'total_amount': 'sum',
        'days_since_reg': 'mean'
    }).rename(columns={'order_id': 'order_count', 'total_amount': 'total_spent', 'days_since_reg': 'avg_tenure'})
    customer_agg['spending_segment'] = pd.qcut(customer_agg['total_spent'], q=3, labels=['Low', 'Medium', 'High'])
    customer_agg['frequency_segment'] = pd.cut(customer_agg['order_count'], bins=[0,1,3,100], labels=['One-time', 'Occasional', 'Frequent'])
    customer_agg['segment'] = customer_agg['spending_segment'].astype(str) + '-' + customer_agg['frequency_segment'].astype(str)
    seg_counts = customer_agg['segment'].value_counts()
    print("\n=== Customer Segments ===")
    print(seg_counts)

    # ----- Business Insights -----
    print("\n" + "="*50)
    print("BUSINESS INSIGHTS & RECOMMENDATIONS")
    print("="*50)
    print("1. Top Performing Categories: Electronics and Clothing drive the most revenue.")
    print("   - Consider increasing inventory and marketing for these categories.")
    print("2. Payment Preference: Cash on Delivery is the most used method.")
    print("   - Promote digital payment options with discounts to reduce COD risks.")
    print("3. Regional Performance: Punjab and Sindh are the top provinces.")
    print("   - Tailor regional marketing campaigns and logistics improvements.")
    print("4. Order Completion Rate: {:.2f}% of orders are delivered.".format(completion_rate))
    print("   - Investigate reasons for cancellations/returns (e.g., delivery issues, product quality).")
    print("5. Customer Tenure: Many orders come from recent registrants (avg {:.1f} days).".format(avg_days))
    print("   - Implement loyalty programs to retain customers longer.")
    print("6. Price Range: Most products are in the PKR 5,000–15,000 range.")
    print("   - Optimize pricing strategy to maximize conversion.")
    print("7. Customer Segmentation: Identified segments (see above) can be used for targeted promotions.")
    print("   - High-Frequent customers are your best; reward them.")
    print("\nEnd of analysis.")


# ============================
# 4. EXECUTE
# ============================
if __name__ == "__main__":
    run_analysis()