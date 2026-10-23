# Fabric notebook source

# METADATA ********************

# META {
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse": "lh2",
# META       "default_lakehouse_name": "Gold",
# META       "default_lakehouse_workspace_id": "aaaaaaaa-0000-0000-0000-000000000001",
# META       "known_lakehouses": [ { "id": "lh1" }, { "id": "lh2" } ]
# META     }
# META   }
# META }

# CELL ********************

from pyspark.sql.functions import col
src = 'abfss://aaaaaaaa-0000-0000-0000-000000000001@onelake.dfs.fabric.microsoft.com/lh1/Tables/'
# old = 'abfss://x@oldaccount.dfs.core.windows.net/raw'
df = spark.read.table("silver.sales")
df2 = spark.sql("SELECT * FROM gold.dim_date d JOIN silver.sales s ON 1 = 1")
df.write.mode("overwrite").saveAsTable("gold.fact_sales")
notebookutils.notebook.run("Load bronze", 600)
requests.get("https://api.contoso.com/rates")

# METADATA ********************

# META {
# META   "language": "python"
# META }