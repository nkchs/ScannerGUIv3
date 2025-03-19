using System;
using System.Text;
using System.Text.Json;
using System.Net.Http;
using System.Collections.Generic;
using ScannerGUIv3.Definitions;
using ScannerGUIv3.Core;

namespace ScannerGUIv3.Services
{
    public class CardService
    {
        public static string CreateAdaptiveCardFromTemplate(string url, Dictionary<int, Employee> employeeDict)
        {
            // Define the JSON structure using anonymous objects with explicit array typing
            var teamsMessage = new
            {
                type = "message",
                attachments = new object[]
                {
                        new
                        {
                            contentType = "application/vnd.microsoft.card.adaptive",
                            content = new
                            {
                                type = "AdaptiveCard",
                                schema = "http://adaptivecards.io/schemas/adaptive-card.json",
                                version = "1.6",
                                msteams = new { width = "Full" },
                                body = new object[]
                                {
                                    new
                                    {
                                        type = "TextBlock",
                                        size = "Medium",
                                        weight = "Bolder",
                                        text = "Attendance Report"
                                    },
                                    new
                                    {
                                        type = "TextBlock",
                                        text = "Signed In",
                                        wrap = true
                                    },
                                    new
                                    {
                                        type = "Table",
                                        columns = new object[]
                                        {
                                            new { width = 2 },
                                            new { width = 1 },
                                            new { width = 1 },
                                            new { width = 1 }
                                        },
                                        rows = GetEmployeeRows(employeeDict)
                                    }
                                }
                            }
                        }
                }
            };

            // Serialize to JSON using System.Text.Json
            string json = JsonSerializer.Serialize(teamsMessage, new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null  // Ensures property names match exactly as defined
            });

            //Console.WriteLine(json);

            // Send to the URL
            try
            {
                using var client = new HttpClient();
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                var response = client.PostAsync(url, content).Result;
                Console.WriteLine(response);
                response.EnsureSuccessStatusCode();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error posting to Teams: {ex.Message}");
                throw;
            }

            return json;
        }

        private static object[] GetEmployeeRows(Dictionary<int, Employee> employeeDict)
        {
            var rows = new List<object>
                {
                    new
                    {
                        type = "TableRow",
                        cells = new object[]
                        {
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Name" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "ID" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign In" } }
                            },
                            new
                            {
                                type = "TableCell",
                                items = new object[] { new { type = "TextBlock", weight = "Bolder", text = "Sign Out" } }
                            }
                        }
                    }
                };

            foreach (var employee in employeeDict.Values)
            {
                if (employee.SignInTime.HasValue)
                {
                    rows.Add(new
                    {
                        type = "TableRow",
                        cells = new object[]
                        {
                                new
                                {
                                    type = "TableCell",
                                    items = new object[] { new { type = "TextBlock", text = employee.Name } }
                                },
                                new
                                {
                                    type = "TableCell",
                                    items = new object[] { new { type = "TextBlock", text = employee.EmployeeNumber.ToString() } }
                                },
                                new
                                {
                                    type = "TableCell",
                                    items = new object[] { new { type = "TextBlock", text = employee.FormattedSignInTime } }
                                },
                                new
                                {
                                    type = "TableCell",
                                    items = new object[] { new { type = "TextBlock", text = employee.FormattedSignOutTime ?? "No Sign Out" } }
                                }
                        }
                    });
                }
            }

            return rows.ToArray();
        }
    }
}