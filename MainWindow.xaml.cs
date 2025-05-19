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
					MessageBox.Show($"Updated {rows} row(s)", "Update Result");
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
			string DBName = DatabaseName.Text?.Trim();
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
			System.Diagnostics.Debug.WriteLine("ContentGrid_CellEditEnding triggered");
			if (e.EditAction == DataGridEditAction.Commit)
			{
				await UpdateDatabase(e);
			}
		}

		private async void ResultGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
		{
			System.Diagnostics.Debug.WriteLine("ResultGrid_CellEditEnding triggered");
			if (e.EditAction == DataGridEditAction.Commit)
			{
				await UpdateDatabase(e);
			}
		}

		private async Task UpdateDatabase(DataGridCellEditEndingEventArgs e)
		{
			System.Diagnostics.Debug.WriteLine("UpdateDatabase method started");
			if (string.IsNullOrWhiteSpace(CurrentTableName))
			{
				MessageBox.Show("No table selected. Please select a table first.", "Error");
				System.Diagnostics.Debug.WriteLine("No CurrentTableName");
				return;
			}

			try
			{
				System.Diagnostics.Debug.WriteLine("Getting row and column data");
				var row = (DataRowView)e.Row.Item;
				string columnName = e.Column.Header.ToString();
				var newValue = (e.EditingElement as TextBox)?.Text;

				System.Diagnostics.Debug.WriteLine($"Column: {columnName}, New Value: {newValue}");

				StringBuilder rowData = new StringBuilder();
				foreach (DataColumn col in row.Row.Table.Columns)
				{
					rowData.Append($"{col.ColumnName}: {row[col.ColumnName]}, ");
				}
				System.Diagnostics.Debug.WriteLine($"Row Data: {rowData.ToString().TrimEnd(',', ' ')}");

				string primaryKeyColumn = GetPrimaryKeyColumn(CurrentTableName);
				if (string.IsNullOrWhiteSpace(primaryKeyColumn))
				{
					MessageBox.Show("Could not determine the primary key for the table.", "Error");
					System.Diagnostics.Debug.WriteLine("No primary key found");
					return;
				}

				var primaryKeyValue = row[primaryKeyColumn];
				System.Diagnostics.Debug.WriteLine($"Raw Primary Key Value: {primaryKeyValue}, Type: {primaryKeyValue?.GetType().Name ?? "null"}");

				bool isNewRow = primaryKeyValue == DBNull.Value || primaryKeyValue == null || row.Row.RowState == DataRowState.Added;
				if (isNewRow)
				{
					System.Diagnostics.Debug.WriteLine("Detected new row, calling InsertNewRow");
					await InsertNewRow(row);
					return;
				}

				if (!int.TryParse(primaryKeyValue.ToString(), out int pkValue))
				{
					MessageBox.Show("Invalid primary key value.", "Error");
					System.Diagnostics.Debug.WriteLine($"Invalid primary key value: {primaryKeyValue}");
					return;
				}

				System.Diagnostics.Debug.WriteLine($"Primary Key Column: {primaryKeyColumn}, Primary Key Value: {pkValue}");

				using var conn = new SqlConnection(ConnStr);
				await conn.OpenAsync();
				string existsQuery = $"SELECT COUNT(*) FROM [{CurrentTableName}] WHERE [{primaryKeyColumn}] = @PrimaryKeyValue";
				using var existsCmd = new SqlCommand(existsQuery, conn);
				existsCmd.Parameters.AddWithValue("@PrimaryKeyValue", pkValue);
				int rowCount = (int)await existsCmd.ExecuteScalarAsync();
				System.Diagnostics.Debug.WriteLine($"Row count for ID {pkValue}: {rowCount}");
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
						System.Diagnostics.Debug.WriteLine($"Column {columnName} does not allow nulls");
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
				System.Diagnostics.Debug.WriteLine($"Executing query: {debugQuery}");

				int rowsAffected = await cmd.ExecuteNonQueryAsync();
				System.Diagnostics.Debug.WriteLine($"Rows affected: {rowsAffected}");

				if (rowsAffected > 0)
				{
					MessageBox.Show($"Updated {rowsAffected} row(s)", "Update Result");
					System.Diagnostics.Debug.WriteLine("Update successful");
				}
				else
				{
					MessageBox.Show("No rows were updated. Please check the data.", "Update Result");
					System.Diagnostics.Debug.WriteLine("No rows updated");
					return;
				}

				string verifyQuery = $"SELECT [{columnName}] FROM [{CurrentTableName}] WHERE [{primaryKeyColumn}] = @PrimaryKeyValue";
				using var verifyCmd = new SqlCommand(verifyQuery, conn);
				verifyCmd.Parameters.AddWithValue("@PrimaryKeyValue", pkValue);
				var updatedValue = await verifyCmd.ExecuteScalarAsync();
				System.Diagnostics.Debug.WriteLine($"Updated value in database: {updatedValue}");

				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");
				System.Diagnostics.Debug.WriteLine("Grid refreshed");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Failed to update database: {ex.Message}\nStack Trace: {ex.StackTrace}", "Error");
				System.Diagnostics.Debug.WriteLine($"Exception: {ex.Message}\nStack Trace: {ex.StackTrace}");
			}
		}

		private async Task InsertNewRow(DataRowView row)
		{
			try
			{
				System.Diagnostics.Debug.WriteLine("InsertNewRow method started");
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
					System.Diagnostics.Debug.WriteLine($"Parameter @p_{column} = {value}");
				}

				System.Diagnostics.Debug.WriteLine($"Executing INSERT query: {insertQuery}");
				var newId = await cmd.ExecuteScalarAsync();
				System.Diagnostics.Debug.WriteLine($"Inserted row with ID: {newId}");

				MessageBox.Show($"Inserted 1 row with ID {newId}", "Insert Result");

				// Оновлення DataGrid
				await ExecuteScriptAsync($"SELECT * FROM [{CurrentTableName}]");

				// Оновлення DataRowView з новим ID
				if (row.Row.Table.Columns.Contains(primaryKeyColumn))
				{
					row.Row[primaryKeyColumn] = Convert.ToInt32(newId);
					row.Row.AcceptChanges(); // Прийняти зміни, щоб строка більше не вважалася новою
				}

				System.Diagnostics.Debug.WriteLine("Grid refreshed and row updated with new ID");
			}
			catch (Exception ex)
			{
				
				System.Diagnostics.Debug.WriteLine($"Exception in InsertNewRow: {ex.Message}\nStack Trace: {ex.StackTrace}");
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
			mainWindow.Width = 800;
		}
	}
}