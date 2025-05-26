using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlClient;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using System.Windows.Media.Animation;
using Microsoft.Win32;
using System.IO;
using OfficeOpenXml;
using System.Reflection.Metadata;
using System.Linq;
using System.Diagnostics;


namespace DataBase
{
	public partial class MainWindow : Window
	{
		private string? ConnStr;
		private string? CurrentTableName;

		public MainWindow()
		{
			InitializeComponent();
			AppGrid.Visibility = Visibility.Hidden;
			MenuConnectGrid.Visibility = Visibility.Visible;
			BackToTableButton.Visibility = Visibility.Hidden;
			RunQueryButton.Visibility = Visibility.Hidden;
			ResultGrid.Visibility = Visibility.Hidden;
			Query.Visibility = Visibility.Hidden;

			DocumentLList.Items.Add("Employement");
			DocumentLList.Items.Add("Dismissial");
			DocumentLList.Items.Add("Filters");
		}

		private void DisconnectButton_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				ConnStr = null;
				CurrentTableName = null;

				ContentGrid.ItemsSource = null;
				ResultGrid.ItemsSource = null;
				Tables_List.Items.Clear();
				Query.Text = string.Empty;
				DBNameTextBlock.Text = string.Empty;

				AppGrid.Visibility = Visibility.Hidden;
				MenuConnectGrid.Visibility = Visibility.Visible;

			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to disconnect: {ex.Message}\nStack Trace: {ex.StackTrace}", "Error");
			}
		}

		private void Window_MouseDown(object sender, MouseButtonEventArgs e)
		{
			if (e.ButtonState == MouseButtonState.Pressed)
			{
				try
				{
					this.DragMove();
				}
				catch (Exception InvalidOperationException)
				{
					Console.WriteLine(InvalidOperationException.Message);
				}
			}
		}

		private async void Connect_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				await ExecuteScriptAsync(Query.Text);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error");
			}

		}

		private async Task ExecuteScriptAsync(string sqlScript)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			using var conn = new SqlConnection(ConnStr);
			try
			{
				await conn.OpenAsync();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to connect: {ex.Message}", "Connection Error");
				return;
			}

			var batches = Regex.Split(
				sqlScript,
				@"^\s*GO\s*($|\-\-.*$)",
				RegexOptions.Multiline | RegexOptions.IgnoreCase);

			foreach (var batch in batches)
			{
				var sql = batch.Trim();
				if (string.IsNullOrWhiteSpace(sql)) continue;

				using var cmd = new SqlCommand(sql, conn);

				if (sql.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
					|| sql.StartsWith("WITH", StringComparison.OrdinalIgnoreCase))
				{
					await ShowResultAsync(cmd);
				}
				else if (sql.StartsWith("INSERT", StringComparison.OrdinalIgnoreCase))
				{
					var rows = await cmd.ExecuteNonQueryAsync();
					MessageBox.Show($"Inserted {rows} row(s)", "Create Result");
				}
				else if (sql.StartsWith("UPDATE", StringComparison.OrdinalIgnoreCase))
				{
					var rows = await cmd.ExecuteNonQueryAsync();
					//MessageBox.Show($"Updated {rows} row(s)", "Update Result"); //IT WAS SO FCKNG ANNOYING DO NOT UNCOMMENT THIS SHIT
				}
				else if (sql.StartsWith("DELETE", StringComparison.OrdinalIgnoreCase))
				{
					var rows = await cmd.ExecuteNonQueryAsync();
					MessageBox.Show($"Deleted {rows} row(s)", "Delete Result");
				}
				else
				{
					var rows = await cmd.ExecuteNonQueryAsync();
					MessageBox.Show($"{rows} rows affected", "Result");
				}
			}
		}

		private async Task ShowResultAsync(SqlCommand cmd)
		{
			using var reader = await cmd.ExecuteReaderAsync();
			var table = new DataTable();
			table.Load(reader);

			// Clear existing ItemsSource to force refresh
			ContentGrid.ItemsSource = null;
			ResultGrid.ItemsSource = null;

			// Assign new data and force UI refresh
			ContentGrid.ItemsSource = table.DefaultView;
			ResultGrid.ItemsSource = table.DefaultView;

			// Force the DataGrid to refresh its UI
			ContentGrid.Items.Refresh();
			ResultGrid.Items.Refresh();
			ContentGrid.UpdateLayout();
			ResultGrid.UpdateLayout();
		}

		private async void ConnectToDB_Click(object sender, RoutedEventArgs e)
		{
			string DBName = "ScientificSystem2";
			string HostName = ServerName.Text?.Trim();

			if (string.IsNullOrWhiteSpace(DBName) || string.IsNullOrWhiteSpace(HostName))
			{
				MessageBox.Show("Server name and database name cannot be empty.", "Connection Error");
				return;
			}

			ConnStr =
				$"Server={HostName};" +
				$"Database={DBName};" +
				"Trusted_Connection=True;" +
				"Encrypt=False;";

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();
				DBNameTextBlock.Text = DBName;
				await LoadTableNames(conn);
				AppGrid.Visibility = Visibility.Visible;
				MenuConnectGrid.Visibility = Visibility.Hidden;
			}
			catch (SqlException ex)
			{
				MessageBox.Show("Incorrect server name or database name. Please check your input.", "Connection Error");
				ConnStr = null;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"An error occurred: {ex.Message}", "Connection Error");
				ConnStr = null;
			}
		}

		private async Task LoadTableNames(SqlConnection conn)
		{
			try
			{
				Tables_List.Items.Clear();
				string query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";
				using var cmd = new SqlCommand(query, conn);
				using var reader = await cmd.ExecuteReaderAsync();

				while (await reader.ReadAsync())
				{
					string tableName = reader.GetString(0);
					Tables_List.Items.Add(tableName);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to load table names: {ex.Message}", "Error");
			}
		}

		private async void Tables_List_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (Tables_List.SelectedItem == null)
				return;

			string selectedTable = Tables_List.SelectedItem.ToString();
			if (string.IsNullOrWhiteSpace(selectedTable))
			{
				MessageBox.Show("Selected table name is invalid.", "Error");
				return;
			}

			CurrentTableName = selectedTable;
			string query = $"SELECT * FROM [{selectedTable}]";
			Query.Text = query;

			if (selectedTable == "Academic_degree")
			{
				AddColumnButton.Content = $"Add degree";
				DeleteColumnButton.Content = $"Delete degree";
				EditColumnButton.Content = $"Edit degree";
				DeleteColumnButton.FontSize = 12;
				AddColumnButton.FontSize = 12;
				EditColumnButton.FontSize = 12;
			}
			else if (selectedTable == "Academic_title")
			{
				AddColumnButton.Content = $"Add academic title";
				EditColumnButton.Content = $"Edit academic title";
				DeleteColumnButton.FontSize = 11;
				AddColumnButton.FontSize = 11;
				EditColumnButton.FontSize = 11;
			}
			else
			{
				AddColumnButton.Content = $"Add in {selectedTable}";
				DeleteColumnButton.Content = $"Delete in {selectedTable}";
				EditColumnButton.Content = $"Edit in {selectedTable}";
				DeleteColumnButton.FontSize = 12;
				AddColumnButton.FontSize = 12;
				EditColumnButton.FontSize = 12;
			}

			// Заповнення ComboBox списком стовпців таблиці
			FilterColumnComboBox.ItemsSource = GetTableColumns(CurrentTableName);
			FilterColumnComboBox.SelectedIndex = 0; // Вибираємо перший стовпець за замовчуванням

			try
			{
				await ExecuteScriptAsync(query);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to execute query: {ex.Message}", "Error");
			}
		}

		private void Querry_Button_Click(object sender, RoutedEventArgs e)
		{
			ContentGrid.Visibility = Visibility.Hidden;
			Querry_Button.Visibility = Visibility.Hidden;
			RunQueryButton.Visibility = Visibility.Visible;
			Query.Visibility = Visibility.Visible;
			BackToTableButton.Visibility = Visibility.Visible;
			ResultGrid.Visibility = Visibility.Visible;
		}

		private void BackToTable_Button_Click(object sender, RoutedEventArgs e)
		{
			ContentGrid.Visibility = Visibility.Visible;
			Querry_Button.Visibility = Visibility.Visible;
			RunQueryButton.Visibility = Visibility.Hidden;
			Query.Visibility = Visibility.Hidden;
			BackToTableButton.Visibility = Visibility.Hidden;
			ResultGrid.Visibility = Visibility.Hidden;
		}

		private async void ContentGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
		{

			if (e.EditAction == DataGridEditAction.Commit)
			{
				await UpdateDatabase(e);
			}
		}

		private async void ResultGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
		{

			if (e.EditAction == DataGridEditAction.Commit)
			{
				await UpdateDatabase(e);
			}
		}

		private async Task UpdateDatabase(DataGridCellEditEndingEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("No table selected. Please select a table first.", "Error");
				return;
			}

			try
			{
				var row = (DataRowView)e.Row.Item;
				string columnName = e.Column.Header.ToString();
				var newValue = (e.EditingElement as TextBox)?.Text;

				//System.Diagnostics.Debug.WriteLine($"Column: {columnName}, New Value: {newValue}");

				StringBuilder rowData = new StringBuilder();
				foreach (DataColumn col in row.Row.Table.Columns)
				{
					rowData.Append($"{col.ColumnName}: {row[col.ColumnName]}, ");
				}
				//System.Diagnostics.Debug.WriteLine($"Row Data: {rowData.ToString().TrimEnd(',', ' ')}");

				string primaryKeyColumn = GetPrimaryKeyColumn(CurrentTableName);
				if (string.IsNullOrWhiteSpace(primaryKeyColumn))
				{
					MessageBox.Show("Could not determine the primary key for the table.", "Error");
					//System.Diagnostics.Debug.WriteLine("No primary key found");
					return;
				}

				var primaryKeyValue = row[primaryKeyColumn];
				//System.Diagnostics.Debug.WriteLine($"Raw Primary Key Value: {primaryKeyValue}, Type: {primaryKeyValue?.GetType().Name ?? "null"}");

				bool isNewRow = primaryKeyValue == DBNull.Value || primaryKeyValue == null || row.Row.RowState == DataRowState.Added;
				if (isNewRow)
				{
					await InsertNewRow(row);
					return;
				}

				if (!int.TryParse(primaryKeyValue.ToString(), out int pkValue))
				{
					MessageBox.Show("Invalid primary key value.", "Error");
					return;
				}

				//System.Diagnostics.Debug.WriteLine($"Primary Key Column: {primaryKeyColumn}, Primary Key Value: {pkValue}");

				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();
				string existsQuery = $"SELECT COUNT(*) FROM [{CurrentTableName}] WHERE [{primaryKeyColumn}] = @PrimaryKeyValue";
				using var existsCmd = new SqlCommand(existsQuery, conn);
				existsCmd.Parameters.AddWithValue("@PrimaryKeyValue", pkValue);
				int rowCount = (int)await existsCmd.ExecuteScalarAsync();
				//System.Diagnostics.Debug.WriteLine($"Row count for ID {pkValue}: {rowCount}");
				if (rowCount == 0)
				{
					MessageBox.Show($"Press enter again to insert a new row", "Warning");
					return;
				}

				string columnAllowsNulls = GetColumnAllowsNulls(CurrentTableName, columnName);
				object formattedNewValue;
				if (string.IsNullOrWhiteSpace(newValue))
				{
					if (columnAllowsNulls == "YES")
					{
						formattedNewValue = DBNull.Value;
					}
					else
					{
						MessageBox.Show($"Column {columnName} does not allow null values. Please enter a value.", "Error");
						//System.Diagnostics.Debug.WriteLine($"Column {columnName} does not allow nulls");
						return;
					}
				}
				else
				{
					var columnType = GetColumnDataType(CurrentTableName, columnName);
					formattedNewValue = FormatValueForColumn(newValue, columnType);
				}

				string updateQuery = $"UPDATE [{CurrentTableName}] SET [{columnName}] = @NewValue WHERE [{primaryKeyColumn}] = @PrimaryKeyValue";

				using var cmd = new SqlCommand(updateQuery, conn);
				cmd.Parameters.AddWithValue("@NewValue", formattedNewValue);
				cmd.Parameters.AddWithValue("@PrimaryKeyValue", pkValue);

				string debugQuery = $"UPDATE [{CurrentTableName}] SET [{columnName}] = '{formattedNewValue}' WHERE [{primaryKeyColumn}] = '{pkValue}'";
				//System.Diagnostics.Debug.WriteLine($"Executing query: {debugQuery}");

				int rowsAffected = await cmd.ExecuteNonQueryAsync();
				//System.Diagnostics.Debug.WriteLine($"Rows affected: {rowsAffected}");

				if (rowsAffected > 0)
				{
					//MessageBox.Show($"Updated {rowsAffected} row(s)", "Update Result"); //IT WAS SO FCKNG ANNOYING DO NOT UNCOMMENT THIS SHIT
					//System.Diagnostics.Debug.WriteLine("Update successful");
				}
				else
				{
					MessageBox.Show("No rows were updated. Please check the data.", "Update Result");
					//System.Diagnostics.Debug.WriteLine("No rows updated");
					return;
				}

				string verifyQuery = $"SELECT [{columnName}] FROM [{CurrentTableName}] WHERE [{primaryKeyColumn}] = @PrimaryKeyValue";
				using var verifyCmd = new SqlCommand(verifyQuery, conn);
				verifyCmd.Parameters.AddWithValue("@PrimaryKeyValue", pkValue);
				var updatedValue = await verifyCmd.ExecuteScalarAsync();
				//System.Diagnostics.Debug.WriteLine($"Updated value in database: {updatedValue}");

				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");
				//System.Diagnostics.Debug.WriteLine("Grid refreshed");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to update database: {ex.Message}\nStack Trace: {ex.StackTrace}", "Error");
				//System.Diagnostics.Debug.WriteLine($"Exception: {ex.Message}\nStack Trace: {ex.StackTrace}");
			}
		}

		private async Task InsertNewRow(DataRowView row)
		{
			try
			{
				//System.Diagnostics.Debug.WriteLine("InsertNewRow method started");
				string[] columns = GetTableColumns(CurrentTableName);
				string primaryKeyColumn = GetPrimaryKeyColumn(CurrentTableName);
				var insertColumns = columns.Where(c => c != primaryKeyColumn).ToList();

				if (!insertColumns.Any())
				{
					MessageBox.Show("No columns available to insert.", "Error");
					return;
				}

				string columnNames = string.Join(", ", insertColumns.Select(c => $"[{c}]"));
				string parameterNames = string.Join(", ", insertColumns.Select(c => $"@p_{c}"));
				string insertQuery = $"INSERT INTO [{CurrentTableName}] ({columnNames}) VALUES ({parameterNames}); SELECT SCOPE_IDENTITY();";

				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();
				using var cmd = new SqlCommand(insertQuery, conn);

				foreach (var column in insertColumns)
				{
					var value = row[column];
					string columnType = GetColumnDataType(CurrentTableName, column);
					string allowsNulls = GetColumnAllowsNulls(CurrentTableName, column);

					if (value == null || value == DBNull.Value || (value is string str && string.IsNullOrWhiteSpace(str)))
					{
						if (allowsNulls == "YES")
						{
							cmd.Parameters.AddWithValue($"@p_{column}", DBNull.Value);
						}
						else
						{
							MessageBox.Show($"Column {column} does not allow null values. Please enter a value.", "Error");
							return;
						}
					}
					else
					{
						object formattedValue = FormatValueForColumn(value.ToString(), columnType);
						cmd.Parameters.AddWithValue($"@p_{column}", formattedValue);
					}
					//System.Diagnostics.Debug.WriteLine($"Parameter @p_{column} = {value}");
				}

				//System.Diagnostics.Debug.WriteLine($"Executing INSERT query: {insertQuery}");
				var newId = await cmd.ExecuteScalarAsync();
				//System.Diagnostics.Debug.WriteLine($"Inserted row with ID: {newId}");

				MessageBox.Show($"Inserted 1 row with ID {newId}", "Insert Result");


				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");


				if (row.Row.Table.Columns.Contains(primaryKeyColumn))
				{
					row.Row[primaryKeyColumn] = Convert.ToInt32(newId);
					row.Row.AcceptChanges();
				}

				//System.Diagnostics.Debug.WriteLine("Grid refreshed and row updated with new ID");
			}
			catch (Exception ex)
			{

				//System.Diagnostics.Debug.WriteLine($"Exception in InsertNewRow: {ex.Message}\nStack Trace: {ex.StackTrace}");
			}
		}

		private string GetColumnAllowsNulls(string tableName, string columnName)
		{
			try
			{
				using var conn = new SqlConnection(ConnStr);
				conn.Open();
				string query = $@"
					SELECT IS_NULLABLE
					FROM INFORMATION_SCHEMA.COLUMNS
					WHERE TABLE_NAME = @TableName
					AND COLUMN_NAME = @ColumnName";
				using var cmd = new SqlCommand(query, conn);
				cmd.Parameters.AddWithValue("@TableName", tableName);
				cmd.Parameters.AddWithValue("@ColumnName", columnName);
				var result = cmd.ExecuteScalar();
				return result?.ToString() ?? "NO";
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to get column nullability: {ex.Message}", "Error");
				System.Diagnostics.Debug.WriteLine($"Failed to get column nullability: {ex.Message}");
				return "NO";
			}
		}

		private string GetPrimaryKeyColumn(string tableName)
		{
			try
			{
				using var conn = new SqlConnection(ConnStr);
				conn.Open();
				string query = $@"
                    SELECT COLUMN_NAME
                    FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
                    WHERE TABLE_NAME = @TableName
                    AND CONSTRAINT_NAME LIKE 'PK_%'";
				using var cmd = new SqlCommand(query, conn);
				cmd.Parameters.AddWithValue("@TableName", tableName);
				var result = cmd.ExecuteScalar();
				return result?.ToString();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to get primary key: {ex.Message}", "Error");
				System.Diagnostics.Debug.WriteLine($"Failed to get primary key: {ex.Message}");
				return null;
			}
		}

		private string GetColumnDataType(string tableName, string columnName)
		{
			try
			{
				using var conn = new SqlConnection(ConnStr);
				conn.Open();
				string query = $@"
                    SELECT DATA_TYPE
                    FROM INFORMATION_SCHEMA.COLUMNS
                    WHERE TABLE_NAME = @TableName
                    AND COLUMN_NAME = @ColumnName";
				using var cmd = new SqlCommand(query, conn);
				cmd.Parameters.AddWithValue("@TableName", tableName);
				cmd.Parameters.AddWithValue("@ColumnName", columnName);
				var result = cmd.ExecuteScalar();
				return result?.ToString()?.ToLower();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to get column data type: {ex.Message}", "Error");
				System.Diagnostics.Debug.WriteLine($"Failed to get column data type: {ex.Message}");
				return "nvarchar";
			}
		}

		private object FormatValueForColumn(string value, string dataType)
		{
			if (string.IsNullOrWhiteSpace(value))
				return DBNull.Value;

			try
			{
				switch (dataType)
				{
					case "int":
						return int.Parse(value);
					case "bigint":
						return long.Parse(value);
					case "bit":
						return value.ToLower() == "true" || value == "1" ? true : false;
					case "datetime":
						return DateTime.Parse(value);
					case "nvarchar":
					case "varchar":
					case "nchar":
					case "char":
						return value;
					default:
						return value;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to format value for column type '{dataType}': {ex.Message}", "Error");
				System.Diagnostics.Debug.WriteLine($"Failed to format value: {ex.Message}");
				return DBNull.Value;
			}
		}

		private async void AddTableButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			string newTableName = Interaction.InputBox("Enter the name for the new table:", "Add Table", "NewTable");
			if (string.IsNullOrWhiteSpace(newTableName))
			{
				MessageBox.Show("Table name cannot be empty.", "Error");
				return;
			}

			if (!Regex.IsMatch(newTableName, @"^[a-zA-Z_][a-zA-Z0-9_]*$"))
			{
				MessageBox.Show("Invalid table name. Use letters, numbers, and underscores only. Start with a letter or underscore.", "Error");
				return;
			}

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string createQuery = $@"
                    CREATE TABLE [{newTableName}] (
                        ID INT PRIMARY KEY IDENTITY(1,1)
                    )";

				using var cmd = new SqlCommand(createQuery, conn);
				await cmd.ExecuteNonQueryAsync();

				MessageBox.Show($"Table '{newTableName}' created successfully.", "Success");
				await LoadTableNames(conn);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to create table: {ex.Message}", "Error");
			}
		}

		private async void DeleteTableButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			if (Tables_List.SelectedItem == null)
			{
				MessageBox.Show("Please select a table to delete.", "Error");
				return;
			}

			string tableToDelete = Tables_List.SelectedItem.ToString();
			if (string.IsNullOrWhiteSpace(tableToDelete))
			{
				MessageBox.Show("Selected table name is invalid.", "Error");
				return;
			}

			var result = MessageBox.Show($"Are you sure you want to delete the table '{tableToDelete}'? This action cannot be undone.", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);
			if (result != MessageBoxResult.Yes)
				return;

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string dropQuery = $"DROP TABLE [{tableToDelete}]";
				using var cmd = new SqlCommand(dropQuery, conn);
				await cmd.ExecuteNonQueryAsync();

				MessageBox.Show($"Table '{tableToDelete}' deleted successfully.", "Success");

				if (CurrentTableName == tableToDelete)
				{
					ContentGrid.ItemsSource = null;
					ResultGrid.ItemsSource = null;
					CurrentTableName = null;
					Query.Text = string.Empty;
				}

				await LoadTableNames(conn);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to delete table: {ex.Message}", "Error");
			}
		}

		private async void AddColumnButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("Please select a table first.", "Error");
				return;
			}

			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			string columnName = Interaction.InputBox("Enter the name for the new column:", "Add Column", "NewColumn");
			if (string.IsNullOrWhiteSpace(columnName))
			{
				MessageBox.Show("Column name cannot be empty.", "Error");
				return;
			}

			if (!Regex.IsMatch(columnName, @"^[a-zA-Z_][a-zA-Z0-9_]*$"))
			{
				MessageBox.Show("Invalid column name. Use letters, numbers, and underscores only. Start with a letter or underscore.", "Error");
				return;
			}

			string[] dataTypes = { "INT", "NVARCHAR(50)", "DATETIME", "BIT" };
			string dataType = Interaction.InputBox("Select data type (e.g., INT, NVARCHAR(50), DATETIME, BIT):", "Add Column", "NVARCHAR(50)");
			if (!dataTypes.Contains(dataType.ToUpper()))
			{
				MessageBox.Show("Invalid or unsupported data type. Choose from INT, NVARCHAR(50), DATETIME, BIT.", "Error");
				return;
			}

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string addColumnQuery = $"ALTER TABLE [{CurrentTableName}] ADD [{columnName}] {dataType}";
				using var cmd = new SqlCommand(addColumnQuery, conn);
				await cmd.ExecuteNonQueryAsync();

				MessageBox.Show($"Column '{columnName}' added to '{CurrentTableName}' successfully.", "Success");

				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to add column: {ex.Message}", "Error");
			}
		}

		private async void EditColumnButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("Please select a table first.", "Error");
				return;
			}

			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			string[] columns = GetTableColumns(CurrentTableName);
			if (columns.Length == 0)
			{
				MessageBox.Show("No columns available to edit.", "Error");
				return;
			}

			string currentColumn = Interaction.InputBox("Select column to edit (e.g., " + string.Join(", ", columns) + "):", "Edit Column", columns[0]);
			if (!columns.Contains(currentColumn))
			{
				MessageBox.Show("Invalid column name.", "Error");
				return;
			}

			string newColumnName = Interaction.InputBox("Enter new column name (leave blank to keep current name):", "Edit Column", currentColumn);
			if (!string.IsNullOrWhiteSpace(newColumnName) && !Regex.IsMatch(newColumnName, @"^[a-zA-Z_][a-zA-Z0-9_]*$"))
			{
				MessageBox.Show("Invalid column name. Use letters, numbers, and underscores only. Start with a letter or underscore.", "Error");
				return;
			}

			if (newColumnName == currentColumn)
				newColumnName = null;

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				if (!string.IsNullOrWhiteSpace(newColumnName))
				{
					string alterQuery = $"EXEC sp_rename '{CurrentTableName}.{currentColumn}', '{newColumnName}', 'COLUMN'";
					using var cmd = new SqlCommand(alterQuery, conn);
					await cmd.ExecuteNonQueryAsync();
					MessageBox.Show($"Column '{currentColumn}' renamed to '{newColumnName}' successfully.", "Success");
				}
				else
				{
					MessageBox.Show("No changes made to column name.", "Info");
					return;
				}

				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to edit column: {ex.Message}", "Error");
			}
		}

		private async void DeleteColumnButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("Please select a table first.", "Error");
				return;
			}

			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			string[] columns = GetTableColumns(CurrentTableName);
			if (columns.Length == 0)
			{
				MessageBox.Show("No columns available to delete.", "Error");
				return;
			}

			string columnToDelete = Interaction.InputBox("Select column to delete (e.g., " + string.Join(", ", columns) + "):", "Delete Column", columns[0]);
			if (!columns.Contains(columnToDelete))
			{
				MessageBox.Show("Invalid column name.", "Error");
				return;
			}

			if (IsPrimaryKey(CurrentTableName, columnToDelete))
			{
				MessageBox.Show("Cannot delete the primary key column.", "Error");
				return;
			}

			var result = MessageBox.Show($"Are you sure you want to delete the column '{columnToDelete}' from '{CurrentTableName}'? This action cannot be undone.", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);
			if (result != MessageBoxResult.Yes)
				return;

			try
			{
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string dropColumnQuery = $"ALTER TABLE [{CurrentTableName}] DROP COLUMN [{columnToDelete}]";
				using var cmd = new SqlCommand(dropColumnQuery, conn);
				await cmd.ExecuteNonQueryAsync();

				MessageBox.Show($"Column '{columnToDelete}' deleted from '{CurrentTableName}' successfully.", "Success");

				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to delete column: {ex.Message}", "Error");
			}
		}

		private string[] GetTableColumns(string tableName)
		{
			var columns = new List<string>();
			try
			{
				using var conn = new SqlConnection(ConnStr);
				conn.Open();
				string query = $"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @TableName";
				using var cmd = new SqlCommand(query, conn);
				cmd.Parameters.AddWithValue("@TableName", tableName);
				using var reader = cmd.ExecuteReader();
				while (reader.Read())
				{
					columns.Add(reader.GetString(0));
				}
			}
			catch (Exception)
			{
				// Handle silently or log if needed
			}
			return columns.ToArray();
		}

		private bool IsPrimaryKey(string tableName, string columnName)
		{
			try
			{
				using var conn = new SqlConnection(ConnStr);
				conn.Open();
				string query = $@"
                    SELECT COUNT(*)
                    FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
                    WHERE TABLE_NAME = @TableName
                    AND COLUMN_NAME = @ColumnName
                    AND CONSTRAINT_NAME LIKE 'PK_%'";
				using var cmd = new SqlCommand(query, conn);
				cmd.Parameters.AddWithValue("@TableName", tableName);
				cmd.Parameters.AddWithValue("@ColumnName", columnName);
				int count = (int)cmd.ExecuteScalar();
				return count > 0;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private void CloseButton_Click(object sender, RoutedEventArgs e)
		{
			Close();
		}

		private void MinimizeButton_Click(object sender, RoutedEventArgs e)
		{
			Window.GetWindow(this).WindowState = WindowState.Minimized;
		}

		private void Button_PreviewMouseDown(object sender, MouseButtonEventArgs e)
		{
			if (sender is Button button && button.RenderTransform is ScaleTransform scaleTransform)
			{
				DoubleAnimation scaleAnim = new DoubleAnimation
				{
					From = 1,
					To = 0.9,
					Duration = TimeSpan.FromSeconds(0.1)
				};

				scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
				scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
			}
		}

		private void Button_PreviewMouseUp(object sender, MouseButtonEventArgs e)
		{
			if (sender is Button button && button.RenderTransform is ScaleTransform scaleTransform)
			{
				DoubleAnimation scaleAnim = new DoubleAnimation
				{
					From = 0.9,
					To = 1,
					Duration = TimeSpan.FromSeconds(0.1)
				};

				scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
				scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
			}
		}

		private void NormalizeButton_Click(object sender, RoutedEventArgs e)
		{
			var mainWindow = Application.Current.MainWindow;

			if (mainWindow.WindowState == WindowState.Maximized)
			{
				mainWindow.WindowState = WindowState.Normal;
			}
			else
			{
				mainWindow.WindowState = WindowState.Maximized;
			}
			mainWindow.Height = 450;
			mainWindow.Width = 1200;
		}

		private async void FilterButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("Please select a table first.", "Error");
				return;
			}

			string filterValue = FilterTextBox.Text.Trim();
			if (string.IsNullOrWhiteSpace(filterValue) || filterValue == "Enter value to filter")
			{
				MessageBox.Show("Please enter a value to filter.", "Error");
				return;
			}

			string selectedColumn = FilterColumnComboBox.SelectedItem?.ToString();
			if (string.IsNullOrWhiteSpace(selectedColumn))
			{
				MessageBox.Show("Please select a column to filter by.", "Error");
				return;
			}

			try
			{
				string filterQuery = $"SELECT * FROM [{CurrentTableName}] WHERE [{selectedColumn}] LIKE @FilterValue";
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();
				using var cmd = new SqlCommand(filterQuery, conn);
				cmd.Parameters.AddWithValue("@FilterValue", $"%{filterValue}%");

				using var reader = await cmd.ExecuteReaderAsync();
				var table = new DataTable();
				table.Load(reader);

				ContentGrid.ItemsSource = null;
				ResultGrid.ItemsSource = null;
				ContentGrid.ItemsSource = table.DefaultView;
				ResultGrid.ItemsSource = table.DefaultView;

				ContentGrid.Items.Refresh();
				ResultGrid.Items.Refresh();
				ContentGrid.UpdateLayout();
				ResultGrid.UpdateLayout();

				MessageBox.Show($"Found {table.Rows.Count} matching record(s).", "Filter Result");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to filter data: {ex.Message}", "Error");
			}
		}
		private void FilterTextBox_GotFocus(object sender, RoutedEventArgs e)
		{
			if (FilterTextBox.Text == "Enter document name")
			{
				FilterTextBox.Text = string.Empty;
				FilterTextBox.Foreground = Brushes.Black;
			}
		}

		private void FilterTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(FilterTextBox.Text))
			{
				FilterTextBox.Text = "Enter document name";
				FilterTextBox.Foreground = Brushes.Gray;
			}
		}

		private async void SaveButton_Click(object sender, RoutedEventArgs e)
		{
			if (ContentGrid.ItemsSource == null || ContentGrid.Items.Count == 0)
			{
				MessageBox.Show("No data to save.", "Error");
				return;
			}

			try
			{
				// Отримуємо дані з ContentGrid
				var dataView = ContentGrid.ItemsSource as DataView;
				if (dataView == null || dataView.Table == null)
				{
					MessageBox.Show("Failed to retrieve data to save.", "Error");
					return;
				}

				DataTable dataTable = dataView.Table;

				// Визначаємо шлях до папки Docs/Filters
				string baseDir = AppDomain.CurrentDomain.BaseDirectory;
				string docsFolder = System.IO.Path.Combine(baseDir, "Docs", "Filters");

				// Створюємо папку, якщо вона не існує
				Directory.CreateDirectory(docsFolder);

				// Збереження у CSV
				StringBuilder csvContent = new StringBuilder();
				csvContent.AppendLine($"Report: Data Export from {CurrentTableName}");
				csvContent.AppendLine($"Generated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
				string filterValue = FilterTextBox.Text.Trim();
				string selectedColumn = FilterColumnComboBox.SelectedItem?.ToString();
				if (!string.IsNullOrWhiteSpace(filterValue) && filterValue != "Enter value to filter" && !string.IsNullOrWhiteSpace(selectedColumn))
				{
					csvContent.AppendLine($"Filter Parameters: {selectedColumn} contains \"{filterValue}\"");
				}
				else
				{
					csvContent.AppendLine("Filter Parameters: None");
				}
				csvContent.AppendLine();

				IEnumerable<string> columnNames = dataTable.Columns.Cast<DataColumn>().Select(column => $"\"{column.ColumnName}\"");
				csvContent.AppendLine(string.Join(",", columnNames));
				foreach (DataRow row in dataTable.Rows)
				{
					IEnumerable<string> fields = row.ItemArray.Select(field =>
					{
						string escapedField = field?.ToString() ?? "";
						escapedField = $"\"{escapedField.Replace("\"", "\"\"")}\"";
						return escapedField;
					});
					csvContent.AppendLine(string.Join(",", fields));
				}

				// Insert into Document table and retrieve the new document_id
				string insertDocumentQuery = @"
            INSERT INTO dbo.Document (Document_type, Employee_id, Document_date, Description)
            OUTPUT INSERTED.Document_id
            VALUES (@DocumentType, @EmployeeId, @DocumentDate, @Description)";

				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				using var docCmd = new SqlCommand(insertDocumentQuery, conn);
				docCmd.Parameters.AddWithValue("@EmployeeId", 0);
				docCmd.Parameters.AddWithValue("@DocumentDate", DateTime.Now);
				docCmd.Parameters.AddWithValue("@DocumentType", "Filters");
				docCmd.Parameters.AddWithValue("@Description", "generated automatically");

				int documentId = (int)await docCmd.ExecuteScalarAsync();
				string csvFilePath = System.IO.Path.Combine(docsFolder, $"_report_{documentId}.csv");

				File.WriteAllText(csvFilePath, csvContent.ToString(), Encoding.UTF8);

				MessageBox.Show($"Report saved successfully to {csvFilePath}", "Success");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to save results: {ex.Message}", "Error");
			}
		}

		private async void Document_List_MouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			CurrentTableName = "Document";
			string query = "";
			string documentType = DocumentLList.SelectedItem.ToString();

			switch (documentType)
			{
				case "Employement":
					query = $"SELECT * FROM dbo.Document WHERE Document_type = 'Employement'";
					Query.Text = query;
					break;
				case "Dismissial":
					query = $"SELECT * FROM dbo.Document WHERE Document_type = 'Dismissial'";
					Query.Text = query;
					break;
				case "Filters":
					query = $"SELECT * FROM dbo.Document WHERE Document_type = 'Filters'";
					Query.Text = query;
					break;
				default:
					query = "";
					MessageBox.Show("Unknown document type selected.", "Error");
					return;
			}

			try
			{
				await ExecuteScriptAsync(query);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to execute query: {ex.Message}", "Error");
			}
		}

		private void OpenDocument(string docsFolder, string filename)
		{
			string fullPath = System.IO.Path.Combine(docsFolder, filename);


			if (!File.Exists(fullPath))
			{
				MessageBox.Show($"File not found: {fullPath}");
				return;
			}

			ProcessStartInfo startInfo = new ProcessStartInfo
			{
				FileName = fullPath,
				UseShellExecute = true,
				WindowStyle = ProcessWindowStyle.Normal
			};

			Process.Start(startInfo);
		}

		private async void CellDoubleClick(object sender, MouseButtonEventArgs e)
		{
			if (CurrentTableName == "Document")
			{
				try
				{
					// Ensure the sender is a DataGrid
					if (!(sender is DataGrid dataGrid))
					{
						MessageBox.Show("Invalid sender: not a DataGrid.");
						return;
					}

					// Get the original source as a DependencyObject
					var originalSource = e.OriginalSource as DependencyObject;
					if (originalSource == null)
					{
						MessageBox.Show("Invalid source: not a DependencyObject.");
						return;
					}

					// Try to find the DataGridCell
					var cell = FindVisualParent<DataGridCell>(originalSource);
					if (cell == null)
					{
						// Debug: Show what was clicked
						string clickedType = originalSource.GetType().Name;


						// Alternative: Try hit-testing to find the cell
						cell = GetCellFromPoint(dataGrid, e.GetPosition(dataGrid));
						if (cell == null)
						{

							return;
						}
					}

					// Get the DataGridRow from the cell
					var row = FindVisualParent<DataGridRow>(cell);
					if (row == null)
					{
						MessageBox.Show("No row was found for the clicked cell.");
						return;
					}

					// Get the data item associated with the row
					var dataItem = row.Item;
					if (dataItem == null || dataItem == CollectionView.NewItemPlaceholder)
					{
						MessageBox.Show("No valid data item associated with the row.");
						return;
					}

					// Cast to DataRowView and access the column directly
					if (!(dataItem is DataRowView dataRowView))
					{
						MessageBox.Show("Data item is not a DataRowView.");
						return;
					}

					// Access the Documentid column by name
					object idValue, personId;
					try
					{
						idValue = dataRowView["Document_id"];
						personId = dataRowView["Employee_id"];

					}
					catch (Exception ex)
					{
						MessageBox.Show($"Error accessing Documentid: {ex.Message}");
						return;
					}

					//MessageBox.Show($"Clicked ID: {idValue?.ToString() ?? "null"}");


					string docsFolder;
					string baseDir = AppDomain.CurrentDomain.BaseDirectory;

					string documentType = DocumentLList.SelectedItem?.ToString();
					if (string.IsNullOrEmpty(documentType))
					{
						MessageBox.Show("No document type selected.", "Error");
						return;
					}

					if (!int.TryParse(idValue.ToString(), out int documentId))
					{
						MessageBox.Show("Invalid Document_id format.");
						return;
					}

					string filename = $"{personId}.txt";

					//MessageBox.Show($"Opening document: {filename}");

					switch (documentType)
					{
						case "Employement":
							docsFolder = System.IO.Path.Combine(baseDir, "Docs", "Hired");
							OpenDocument(docsFolder, filename);
							return;
						case "Dismissial":
							docsFolder = System.IO.Path.Combine(baseDir, "Docs", "Fired");
							OpenDocument(docsFolder, filename);
							return;
						case "Filters":
							docsFolder = System.IO.Path.Combine(baseDir, "Docs", "Filters");
							filename = $"_report_{documentId}.csv";
							OpenDocument(docsFolder, filename);
							return;
						default:
							MessageBox.Show("Unknown document type selected.", "Error");
							return;
					}

				}
				catch (Exception ex)
				{
					MessageBox.Show($"An error occurred: {ex.Message}");
				}
			}
		}



		// Helper method to find a visual parent of type T
		private T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
		{
			while (child != null && !(child is T))
			{
				child = VisualTreeHelper.GetParent(child);
			}
			return child as T;
		}

		// Helper method to find a DataGridCell using hit-testing
		private DataGridCell GetCellFromPoint(DataGrid dataGrid, Point clickPosition)
		{
			// Perform hit-test to find the visual element at the click position
			var hitTestResult = VisualTreeHelper.HitTest(dataGrid, clickPosition);
			if (hitTestResult?.VisualHit == null)
				return null;

			// Traverse up to find the DataGridCell
			return FindVisualParent<DataGridCell>(hitTestResult.VisualHit);
		}

		private async void HireButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			try
			{
				// Prompt for person details
				string firstName = Interaction.InputBox("Enter First Name:", "New Person", "");
				if (string.IsNullOrWhiteSpace(firstName))
				{
					MessageBox.Show("First name cannot be empty.", "Error");
					return;
				}

				string lastName = Interaction.InputBox("Enter Last Name:", "New Person", "");
				if (string.IsNullOrWhiteSpace(lastName))
				{
					MessageBox.Show("Last name cannot be empty.", "Error");
					return;
				}

				string birthDate = Interaction.InputBox("Enter Birth Date (yyyy-MM-dd):", "New Person", "1990-01-01");
				if (!DateTime.TryParse(birthDate, out DateTime parsedBirthDate))
				{
					MessageBox.Show("Invalid birth date format. Please use yyyy-MM-dd.", "Error");
					return;
				}

				string passport = Interaction.InputBox("Enter Passport Number:", "New Person", "");
				if (string.IsNullOrWhiteSpace(passport))
				{
					MessageBox.Show("Passport number cannot be empty.", "Error");
					return;
				}

				string adress = Interaction.InputBox("Enter Address:", "New Person", "");
				if (string.IsNullOrWhiteSpace(adress))
				{
					MessageBox.Show("Address cannot be empty.", "Error");
					return;
				}

				string email = Interaction.InputBox("Enter Email:", "New Person", "");
				if (string.IsNullOrWhiteSpace(email))
				{
					MessageBox.Show("Email cannot be empty.", "Error");
					return;
				}

				string phone = Interaction.InputBox("Enter Phone Number:", "New Person", "");
				if (string.IsNullOrWhiteSpace(phone))
				{
					MessageBox.Show("Phone number cannot be empty.", "Error");
					return;
				}

				// Prompt for employee details
				string jobTitle = Interaction.InputBox("Enter Job Title:", "New Employee", "");
				if (string.IsNullOrWhiteSpace(jobTitle))
				{
					MessageBox.Show("Job title cannot be empty.", "Error");
					return;
				}

				string hireDate = Interaction.InputBox("Enter Hire Date (yyyy-MM-dd):", "New Employee", DateTime.Now.ToString("yyyy-MM-dd"));
				if (!DateTime.TryParse(hireDate, out DateTime parsedHireDate))
				{
					MessageBox.Show("Invalid hire date format. Please use yyyy-MM-dd.", "Error");
					return;
				}

				string departmentIdInput = Interaction.InputBox("Enter Department ID (e.g., 1-8):", "New Employee", "1");
				if (!int.TryParse(departmentIdInput, out int departmentId) || departmentId < 1 || departmentId > 8)
				{
					MessageBox.Show("Invalid Department ID. Please enter a number between 1 and 8.", "Error");
					return;
				}

				string positionIdInput = Interaction.InputBox("Enter Position ID (e.g., 1-5):", "New Employee", "1");
				if (!int.TryParse(positionIdInput, out int positionId) || positionId < 1 || positionId > 5)
				{
					MessageBox.Show("Invalid Position ID. Please enter a number between 1 and 5.", "Error");
					return;
				}

				string salaryInput = Interaction.InputBox("Enter Annual Salary (£):", "New Employee", "30000");
				if (!decimal.TryParse(salaryInput, out decimal salary) || salary <= 0)
				{
					MessageBox.Show("Invalid salary amount.", "Error");
					return;
				}

				// Insert into Person table
				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string insertPersonQuery = @"
            INSERT INTO Person (First_name, Last_name, Birth_date, Passport, Adress, Email, Phone)
            OUTPUT INSERTED.Person_id
            VALUES (@FirstName, @LastName, @BirthDate, @Passport, @Adress, @Email, @Phone)";

				using var personCmd = new SqlCommand(insertPersonQuery, conn);
				personCmd.Parameters.AddWithValue("@FirstName", firstName);
				personCmd.Parameters.AddWithValue("@LastName", lastName);
				personCmd.Parameters.AddWithValue("@BirthDate", parsedBirthDate);
				personCmd.Parameters.AddWithValue("@Passport", passport);
				personCmd.Parameters.AddWithValue("@Adress", adress);
				personCmd.Parameters.AddWithValue("@Email", email);
				personCmd.Parameters.AddWithValue("@Phone", phone);

				var personId = await personCmd.ExecuteScalarAsync();
				if (personId == null)
				{
					MessageBox.Show("Failed to insert person.", "Error");
					return;
				}

				// Insert into Employee table
				string insertEmployeeQuery = @"
					INSERT INTO Employee (Person_id, Department_id, Position_id, Hire_date, Academic_degree_id, Academic_title_id)
					OUTPUT INSERTED.Employee_id
					VALUES (@PersonId, @DepartmentId, @PositionId, @HireDate, @AcademicDegreeId, @AcademicTitleId)";

				using var employeeCmd = new SqlCommand(insertEmployeeQuery, conn);
				employeeCmd.Parameters.AddWithValue("@PersonId", personId);
				employeeCmd.Parameters.AddWithValue("@DepartmentId", departmentId);
				employeeCmd.Parameters.AddWithValue("@PositionId", positionId);
				employeeCmd.Parameters.AddWithValue("@HireDate", parsedHireDate);
				employeeCmd.Parameters.AddWithValue("@AcademicDegreeId", 1);
				employeeCmd.Parameters.AddWithValue("@AcademicTitleId", 1);


				var employeeId = await employeeCmd.ExecuteScalarAsync();
				if (employeeId == null)
				{
					MessageBox.Show("Failed to insert employee.", "Error");
					return;
				}

				// Insert into Document table
				string insertDocumentQuery = @"
					INSERT INTO dbo.Document (Document_type, Employee_id, Document_date, Description)
					VALUES ('Employement', @EmployeeId, @DocumentDate, @Description)";

				using var docCmd = new SqlCommand(insertDocumentQuery, conn);
				docCmd.Parameters.AddWithValue("@EmployeeId", employeeId);
				docCmd.Parameters.AddWithValue("@DocumentDate", DateTime.Now);
				docCmd.Parameters.AddWithValue("@Description", "generated automatically");

				await docCmd.ExecuteNonQueryAsync();

				// Generate and save the contract
				string employerName = "ScientificSystem Ltd"; // Replace with actual employer name or prompt
				string employerAddress = "123 Business Park, Scotland"; // Replace with actual address or prompt
				string registrationNumber = "SC123456"; // Replace with actual registration number
				string docsFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Docs", "Hired");
				string filename = $"{employeeId}.txt";

				try
				{
					if (!Directory.Exists(docsFolder))
					{
						Directory.CreateDirectory(docsFolder);
					}

					string contractContent = GenerateEmployeeContract(
						employerName,
						$"{firstName} {lastName}",
						employerAddress,
						registrationNumber,
						jobTitle,
						parsedHireDate,
						salary,
						adress
					);

					File.WriteAllText(System.IO.Path.Combine(docsFolder, filename), contractContent, Encoding.UTF8);
					MessageBox.Show($"Person and employee added successfully. Person ID: {personId}, Employee ID: {employeeId}. Contract saved to {filename}.", "Success");

					// Refresh the Person or Employee table if currently selected
					if (CurrentTableName == "Person")
					{
						await ExecuteScriptAsync("SELECT * FROM Person");
					}
					else if (CurrentTableName == "Employee")
					{
						await ExecuteScriptAsync("SELECT * FROM Employee");
					}

					// Open the generated document
					OpenDocument(docsFolder, filename);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Failed to save contract: {ex.Message}", "Error");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to add person or employee: {ex.Message}", "Error");
			}
		}

		#region generate employee contract

		private string GenerateEmployeeContract(
			string employerName,
			string employeeName,
			string employerAddress,
			string registrationNumber,
			string jobTitle,
			DateTime startDate,
			decimal salary,
			string employeeAddress)
		{
			StringBuilder contract = new StringBuilder();

			contract.AppendLine($"Employee Contract");
			contract.AppendLine();
			contract.AppendLine($"{employerName}");
			contract.AppendLine();
			contract.AppendLine($"Statement of Terms and Conditions of Employment");
			contract.AppendLine();
			contract.AppendLine($"{employeeName}");
			contract.AppendLine($"{DateTime.Now:yyyy-MM-dd}");
			contract.AppendLine();
			contract.AppendLine($"TERMS AND CONDITIONS OF EMPLOYMENT");
			contract.AppendLine();
			contract.AppendLine($"BETWEEN");
			contract.AppendLine($"(a) {employerName}, an organisation registered in Scotland under registration number {registrationNumber} whose registered office is at {employerAddress} (hereinafter referred to as “the Employer”)");
			contract.AppendLine($"(b) {employeeName} of {employeeAddress} (hereinafter referred to as “you”)");
			contract.AppendLine();
			contract.AppendLine($"IT IS AGREED as follows:");
			contract.AppendLine();
			contract.AppendLine($"1. General");
			contract.AppendLine($"This document is a statement of the main terms and conditions of employment which govern your service with the Employer. Your service with the Employer is also subject to the terms contained in the letter offering you employment ‘the offer letter’. If there should be any ambiguity or discrepancy between the terms in the offer letter and in the terms set out in this document, the terms of the Offer letter will prevail, except where expressly stated to the contrary.");
			contract.AppendLine();
			contract.AppendLine($"2. Duties and Job Title");
			contract.AppendLine($"2.1 You are employed by the Employer in the capacity of {jobTitle}. You will be required to undertake such duties and responsibilities as may be determined by the Employer from time to time.");
			contract.AppendLine($"2.2 The Employer reserves the right to vary your duties and responsibilities at any time and from time to time according to the needs of the Employer’s business, following discussion and agreement with you.");
			contract.AppendLine();
			contract.AppendLine($"3. Date of Commencement/ Date of Continuous employment [and Notice Period]");
			contract.AppendLine($"3.1 Your period of continuous employment with us begins on {startDate:yyyy-MM-dd}.");
			contract.AppendLine($"3.2 No employment with a previous employer counts as part of your period of continuous employment.");
			contract.AppendLine($"3.3 The first 6 months of your employment will be a probationary period. During this period your performance and conduct will be monitored. At the end of the probationary period your performance will be reviewed and if found satisfactory your appointment will be confirmed. The probationary period may be extended at the Employer’s discretion.");
			contract.AppendLine($"During the 6 months probationary period the notice required by either party to this Contract to terminate your employment will be one week.");
			contract.AppendLine();
			contract.AppendLine($"4. Hours of work");
			contract.AppendLine($"4.1 Your normal working hours are between 9:00 am and 5:00 pm Mondays to Fridays inclusive with one hour for lunch which must be taken between 12:00 pm and 2:00 pm.");
			contract.AppendLine($"4.2 The Employer reserves the right to alter working hours as necessary, following discussion and agreement with you.");
			contract.AppendLine($"4.3 You may be asked to work additional hours beyond your normal hours and it is a condition of your employment that you agree to do so when reasonably asked. You will not be entitled to overtime payments for hours worked outside your normal working hours.");
			contract.AppendLine();
			contract.AppendLine($"5. Place of work");
			contract.AppendLine($"Your normal place of work will be at {employerAddress} or such other places as the Employer may reasonably require.");
			contract.AppendLine();
			contract.AppendLine($"6. Remuneration");
			contract.AppendLine($"6.1 Your salary is £{salary:N2} per year, to be paid monthly normally on the last Friday of each month. Payment will be made by direct credit transfer to a bank or building society account nominated by you. You will not be entitled to overtime payment for hours worked outside your normal weekly hours.");
			contract.AppendLine($"6.2 Your salary will be reviewed annually entirely at our discretion.");
			contract.AppendLine($"6.3 The Employer reserves the right to seek reimbursement by deduction from your salary, in accordance with the provisions of the Employment Rights Act 1966 in the event of any material deficiencies attributable to you, in particular damage to Employer property or in the event of overpayment of salary, recovery of unearned holiday pay or other remunerations, or if any other sums are due by you to the Employer arising from your employment.");
			contract.AppendLine();
			contract.AppendLine($"7. Collective agreements");
			contract.AppendLine($"There are no collective agreements relevant to your employment.");
			contract.AppendLine();
			contract.AppendLine($"8. Holidays");
			contract.AppendLine($"8.1 You are entitled to 28 days holiday in each complete calendar year, including bank and public holidays.");
			contract.AppendLine($"8.2 The holiday year commences on January 1 and finishes on December 31 each year.");
			contract.AppendLine($"8.3 If your employment commences or finishes part way through the holiday year, your holiday entitlement will be prorated accordingly.");
			contract.AppendLine($"8.4 If, on termination of employment:");
			contract.AppendLine($"8.4.1 you have exceeded your prorated holiday entitlement, the Employer will deduct a payment in lieu of days holiday taken in excess of your prorated holiday entitlement, on the basis of 1/260th, and you authorise the Employer to make a deduction from the payment of any final salary.");
			contract.AppendLine($"8.4.2 you have holiday entitlement still owing, the Employer may, at its discretion, require you to take your holiday during your notice period or make a payment in lieu of untaken holiday entitlement.");
			contract.AppendLine($"8.5 Holidays must be taken at times convenient to the Employer. You must obtain approval of proposed holiday dates in advance from your Manager. You will not be allowed to take more than two weeks at any one time, save at the Employer’s discretion. You must not book holidays until your request for approval has been formally agreed.");
			contract.AppendLine($"8.6 All holidays must be taken in the year in which it is accrued. In exceptional circumstances you may carry forward up to 5 days untaken holiday entitlement to the next holiday year. This applies for one year only, and holidays may not be carried forward to a subsequent holiday year.");
			contract.AppendLine($"8.7 If you are sick or injured while on holiday, the Employer will allow you to transfer to sick leave and take replacement holiday at a later date, subject to notification and certification requirements.");
			contract.AppendLine();
			contract.AppendLine($"9. Sickness Absence");
			contract.AppendLine($"9.1 In the event of your absence for any reason you or someone on your behalf should contact your Manager at the earliest opportunity and no later than 9:00 am on the first day of the absence to inform him/her of the reason for absence. You must inform the Employer as soon as possible of any change in the date of your expected return to work.");
			contract.AppendLine($"9.2 A self-certification form should be completed for absences of up to seven days. The form will be supplied to you.");
			contract.AppendLine($"9.3 For periods of sickness of more than seven consecutive days, including weekends, you will be required to obtain a Statement of Fitness for Work (‘Fit Note’) / Medical Certificate and send this to your Manager.");
			contract.AppendLine($"9.4 If you are absent for four or more days by reason of sickness or incapacity, you are entitled to Statutory Sick Pay (SSP), provided that you have met the requirements above. For the purposes of the SSP scheme the ‘qualifying days’ are Monday to Friday. There is no contractual right to payment in respect of periods of absence due to sickness or incapacity. Any such payments are at the discretion of the Employer.");
			contract.AppendLine($"9.5 The Employer has the right to monitor and record absence levels and reasons for absences. Such information will be kept confidential.");
			contract.AppendLine($"9.6 The Employer may require you to undergo a medical examination by a medical practitioner nominated by us at any stage of your employment, and you agree to authorise such medical practitioner to prepare a report detailing the results of the examination, which you agree may be disclosed to the Employer. The Employer will bear the cost of such medical examination.");
			contract.AppendLine();
			contract.AppendLine($"10. Maternity and Paternity Rights");
			contract.AppendLine($"The Employer will comply with its statutory obligations with respect to maternity and paternity rights and rights dealing with time off for dependants. The Employer’s policies in this regard are available on request from your Manager.");
			contract.AppendLine();
			contract.AppendLine($"11. Pension");
			contract.AppendLine($"There are no pension arrangements applicable to your employment.");
			contract.AppendLine();
			contract.AppendLine($"12. Non – Compulsory Retirement");
			contract.AppendLine($"The Employer does not operate a normal retirement age and therefore you will not be compulsorily retired on reaching a particular age. However, you can choose to retire voluntarily at any time, provided that you give the required period of notice to terminate your employment.");
			contract.AppendLine();
			contract.AppendLine($"13. Restrictions and Confidentiality");
			contract.AppendLine($"13.1 You may not, without the prior written consent of the Employer, devote any time to any business other than the business of the Employer or to any public or charitable duty or endeavour during your normal hours of work.");
			contract.AppendLine($"13.2 You will not at any time either during your employment or afterwards use or divulge to any person, firm or Employer, except in the proper course of your duties during your employment by the Employer, any confidential information identifying or relating to the Employer, details of which are not in the public domain.");
			contract.AppendLine();
			contract.AppendLine($"14. Mobility");
			contract.AppendLine($"You may be required to travel on Employer business anywhere in the UK. Travel and subsistence will be paid to you in accordance with the Employer’s Expenses Policy.");
			contract.AppendLine();
			contract.AppendLine($"15. Grievance Procedure");
			contract.AppendLine($"The formal Grievance Procedure is available on request from your Manager.");
			contract.AppendLine();
			contract.AppendLine($"16. Disciplinary Procedure");
			contract.AppendLine($"The disciplinary rules applicable to your employment are set out in the attached Disciplinary Rules and Procedure.");
			contract.AppendLine();
			contract.AppendLine($"17. Employee Handbook and Employment Policies");
			contract.AppendLine($"All employees have a duty to adhere to the Employer’s other policies in force, including but not exclusive to the Employer’s Health and Safety, Fire Safety, Sickness and Absence and Equal Opportunities Policies.");
			contract.AppendLine();
			contract.AppendLine($"18. Termination of employment");
			contract.AppendLine($"18.1 During the 6 months probationary period the notice required by either party to this Contract to terminate your employment will be one week.");
			contract.AppendLine($"After the successful completion of any probationary period, your employment may be ended by you giving the Employer one month’s written notice. The Employer will give you one month’s written notice and after four years’ continuous service a further one week’s notice for each additional complete year of service up to a maximum of 12 weeks’ notice.");
			contract.AppendLine($"18.2 The Employer reserves the right in their absolute discretion to pay you salary in lieu of notice.");
			contract.AppendLine($"18.3 Nothing in this Contract prevents the Employer from terminating your employment summarily or otherwise in the event of any serious breach by you of the terms of your employment or in the event of any act or acts of gross misconduct by you.");
			contract.AppendLine();
			contract.AppendLine($"19. Data Protection");
			contract.AppendLine($"You agree to the Employer holding and processing, both electronically and manually, personal data about you (including sensitive personal data as defined in the Data Protection Act 1998) for the operations, management, security or administration of the Employer and for the purpose of complying with applicable laws, regulations and procedures.");
			contract.AppendLine();
			contract.AppendLine($"20. Confidential Information");
			contract.AppendLine($"You will not at any time either during your employment or afterwards use or divulge to any person, firm or Employer, except in the proper course of your duties during your employment by the Employer, any confidential information identifying or relating to the Employer, details of which are not in the public domain.");
			contract.AppendLine();
			contract.AppendLine($"21. Copyright, Inventions and Patents");
			contract.AppendLine($"21.1 All records, documents, papers (including copies and summaries thereof) and other copyright protected works made or acquired by you in the course of your employment shall, together with all the world-wide copyright and design rights in all such works, be and at all times remain the absolute property of the Employer.");
			contract.AppendLine($"21.2 You hereby irrevocably and unconditionally waive all rights granted by Chapter IV of Part I of the Copyright, Designs and Patents Act 1988 that vest in you (whether before, on or after the date hereof) in connection with your authorship of any copyright works in the course of your employment with the Employer, wherever in the world enforceable, including without limitation the right to be identified as the author of any such works and the right not to have any such works subjected to derogatory treatment.");
			contract.AppendLine();
			contract.AppendLine($"22. Changes to Terms and Conditions of Employment");
			contract.AppendLine($"The Employer may amend, vary or terminate the terms and conditions in this document. Any such change to your terms and conditions will be subject to consultation and agreement with you and notified to you personally in writing.");
			contract.AppendLine();
			contract.AppendLine($"23. Severability");
			contract.AppendLine($"The various provision of this Agreement are severable, and if any provision or identifiable part thereof is held to be invalid or unenforceable by any court of competent jurisdiction then such invalidity or unenforceability shall not affect the validity or enforceability of the remaining provisions or identifiable parts.");
			contract.AppendLine();
			contract.AppendLine($"24. Jurisdiction");
			contract.AppendLine($"This Agreement shall be governed by and construed in accordance with Scots Law and Scottish Courts.");
			contract.AppendLine();
			contract.AppendLine($"Issued for and on behalf of {employerName}");
			contract.AppendLine($"Signed: ______________________________ Date: {DateTime.Now:yyyy-MM-dd}");
			contract.AppendLine();
			contract.AppendLine($"Employee");
			contract.AppendLine($"I hereby warrant and confirm that I am not prevented by previous employment terms and conditions, or in any other way, from entering into employment with the Employer or performing any of the duties of employment referred to above. I accept the terms of this Agreement.");
			contract.AppendLine($"Signed: ______________________________ Date: {DateTime.Now:yyyy-MM-dd}");
			contract.AppendLine();
			contract.AppendLine($"{employeeName}");

			return contract.ToString();
		}
		#endregion

		private async void DismissButton_Click(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(ConnStr))
			{
				MessageBox.Show("Please connect to a database first.", "Connection Error");
				return;
			}

			try
			{
				string employeeIdInput = Interaction.InputBox("Enter Employee ID to dismiss:", "Dismiss Employee", "");
				if (!int.TryParse(employeeIdInput, out int employeeId) || employeeId <= 0)
				{
					MessageBox.Show("Invalid Employee ID. Please enter a valid number.", "Error");
					return;
				}

				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();

				string personQuery = @"
            SELECT p.First_name, p.Last_name, p.Adress, e.Hire_date, e.Person_id 
            FROM Employee e 
            JOIN Person p ON e.Person_id = p.Person_id 
            WHERE e.Employee_id = @EmployeeId";

				using var personCmd = new SqlCommand(personQuery, conn);
				personCmd.Parameters.AddWithValue("@EmployeeId", employeeId);

				using var reader = await personCmd.ExecuteReaderAsync();
				if (!reader.HasRows)
				{
					MessageBox.Show("No employee found with the provided ID.", "Error");
					return;
				}

				string firstName = "", lastName = "", employeeAddress = "";
				DateTime hireDate = DateTime.Now;
				int personId = 0;

				if (await reader.ReadAsync())
				{
					firstName = reader.GetString(0);
					lastName = reader.GetString(1);
					employeeAddress = reader.GetString(2);
					hireDate = reader.GetDateTime(3);
					personId = reader.GetInt32(4);
				}
				reader.Close();

				string insertDocumentQuery = @"
            INSERT INTO dbo.Document (Document_type, Employee_id, Document_date, Description)
            OUTPUT INSERTED.Document_id
            VALUES (@DocumentType, @EmployeeId, @DocumentDate, @Description)";

				using var docCmd = new SqlCommand(insertDocumentQuery, conn);
				docCmd.Parameters.AddWithValue("@DocumentType", "Dismissial");
				docCmd.Parameters.AddWithValue("@EmployeeId", employeeId);
				docCmd.Parameters.AddWithValue("@DocumentDate", DateTime.Now);
				docCmd.Parameters.AddWithValue("@Description", "generated automatically");

				int documentId = (int)await docCmd.ExecuteScalarAsync();

				string employerName = "ScientificSystem Ltd";
				string employerAddress = "123 Business Park, Scotland";
				string registrationNumber = "SC123456";
				string docsFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Docs", "Fired");
				string filename = $"{employeeId}.txt";

				try
				{
					if (!Directory.Exists(docsFolder))
					{
						Directory.CreateDirectory(docsFolder);
					}

					string dismissalContent = GenerateDismissalLetter(
						employerName,
						$"{firstName} {lastName}",
						employerAddress,
						registrationNumber,
						hireDate,
						DateTime.Now,
						employeeAddress
					);

					File.WriteAllText(System.IO.Path.Combine(docsFolder, filename), dismissalContent, Encoding.UTF8);

					using var transaction = conn.BeginTransaction();
					try
					{
						string deleteEmployeeQuery = "DELETE FROM Employee WHERE Employee_id = @EmployeeId";
						using var deleteEmpCmd = new SqlCommand(deleteEmployeeQuery, conn, transaction);
						deleteEmpCmd.Parameters.AddWithValue("@EmployeeId", employeeId);
						int empRows = await deleteEmpCmd.ExecuteNonQueryAsync();

						string deletePersonQuery = "DELETE FROM Person WHERE Person_id = @PersonId";
						using var deletePersonCmd = new SqlCommand(deletePersonQuery, conn, transaction);
						deletePersonCmd.Parameters.AddWithValue("@PersonId", personId);
						int personRows = await deletePersonCmd.ExecuteNonQueryAsync();

						transaction.Commit();
						MessageBox.Show(
							$"Employee ID {employeeId} dismissed successfully. {empRows} employee and {personRows} person record(s) deleted. Dismissal letter saved to {filename}.",
							"Success"
						);
					}
					catch (Exception ex)
					{
						transaction.Rollback();
						MessageBox.Show($"Failed to delete records: {ex.Message}", "Error");
						return;
					}

					if (CurrentTableName == "Person")
					{
						await ExecuteScriptAsync("SELECT * FROM Person");
					}
					else if (CurrentTableName == "Employee")
					{
						await ExecuteScriptAsync("SELECT * FROM Employee");
					}

					OpenDocument(docsFolder, filename);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Failed to save dismissal letter: {ex.Message}", "Error");
					return;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to process dismissal: {ex.Message}", "Error");
			}
		}

		private string GenerateDismissalLetter(
			string employerName,
			string employeeName,
			string employerAddress,
			string registrationNumber,
			DateTime hireDate,
			DateTime dismissalDate,
			string employeeAddress)
		{
			StringBuilder letter = new StringBuilder();
			letter.AppendLine($"Dismissal Letter");
			letter.AppendLine();
			letter.AppendLine($"{employerName}");
			letter.AppendLine($"{employerAddress}");
			letter.AppendLine($"Registration Number: {registrationNumber}");
			letter.AppendLine();
			letter.AppendLine($"{employeeName}");
			letter.AppendLine($"{employeeAddress}");
			letter.AppendLine();
			letter.AppendLine($"{DateTime.Now:yyyy-MM-dd}");
			letter.AppendLine();
			letter.AppendLine($"Dear {employeeName},");
			letter.AppendLine();
			letter.AppendLine($"Re: Termination of Employment");
			letter.AppendLine();
			letter.AppendLine($"We regret to inform you that your employment with {employerName} will terminate effective {dismissalDate:yyyy-MM-dd}.");
			letter.AppendLine();
			letter.AppendLine($"You were employed by {employerName} since {hireDate:yyyy-MM-dd}. Following a review of our operational requirements, we have made the difficult decision to terminate your employment.");
			letter.AppendLine();
			letter.AppendLine($"Details of Termination:");
			letter.AppendLine($"- Effective Date: {dismissalDate:yyyy-MM-dd}");
			letter.AppendLine($"- Notice Period: In accordance with your contract, you are entitled to one week's notice, which you will receive as payment in lieu of notice.");
			letter.AppendLine($"- Final Pay: Your final paycheck, including any accrued but unused holiday pay, will be processed and paid on the next scheduled payroll date.");
			letter.AppendLine();
			letter.AppendLine($"Please return any company property in your possession, including but not limited to keys, access cards, and equipment, by {dismissalDate:yyyy-MM-dd}.");
			letter.AppendLine();
			letter.AppendLine($"We thank you for your contributions to {employerName} during your tenure and wish you the best in your future endeavors.");
			letter.AppendLine();
			letter.AppendLine($"Yours sincerely,");
			letter.AppendLine();
			letter.AppendLine($"______________________________");
			letter.AppendLine($"Human Resources Department");
			letter.AppendLine($"{employerName}");

			return letter.ToString();
		}
	}
}