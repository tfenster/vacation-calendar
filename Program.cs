using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System.Text.Json;

// m365 aad app add --name 'Vacation calendar' --redirectUris 'http://localhost/' --platform publicClient --apisDelegated 'https://graph.microsoft.com/Group.ReadWrite.All,https://graph.microsoft.com/Calendars.ReadWrite.Shared,https://graph.microsoft.com/User.Read'

var debugMode = false;
var groupName = "4PS Deutschland";

// 4PS
var clientId = "bc6a5c42-f082-4b55-9a87-e765f30a1ba4";
var tenantId = "92f4dd01-f0ea-4b5f-97f2-505c2945189c";

if (debugMode)
{
    // ars solvendi
    clientId = "43e13dc3-0ca4-4103-a603-5855e988e3c2";
    tenantId = "539f23a3-6819-457e-bd87-7835f4122217";
}
var graphClient = GetGraphClient(clientId, tenantId);

var groupId = await GetGroupId(groupName, graphClient);
if (groupId == null) return;
await CleanCalendar(groupId, graphClient);
var entries = await GetCalendarEntriesFromGroup(groupId, graphClient);

bool initial = true;
foreach (var entry in entries)
{
    if (debugMode && !initial)
    {
        Console.WriteLine($"{entry.Subject} ({entry.Organizer?.EmailAddress?.Name})");
        Console.WriteLine($"\t{entry.Start?.DateTime} - {entry.End?.DateTime}");
        Console.WriteLine();
        continue;
    }
    initial = false;
    await CreateEventInSharedCalendar(entry, graphClient, groupId);
}

static async Task CreateEventInSharedCalendar(Event newEvent, GraphServiceClient graphClient, string groupId)
{
    try
    {
        await graphClient.Groups[groupId].Calendar.Events.PostAsync(new Event()
        {
            Subject = $"{newEvent.Subject} ({newEvent.Organizer?.EmailAddress?.Name})",
            Start = newEvent.Start,
            End = newEvent.End,
        });
    }
    catch (Exception ex)
    {
        HandleException($"Couldn't create event {JsonSerializer.Serialize(newEvent)}", ex);
    }
}

static async Task<List<Event>> GetCalendarEntriesFromGroup(string groupId, GraphServiceClient graphClient)
{
    List<DirectoryObject>? members = await GetMembers(groupId, graphClient);

    var events = new List<Event>();
    foreach (var member in members)
    {
        if (member is User user)
        {
            var userEmail = user.Mail;
            var userId = user.Id;
            try
            {
                var userEvents = await graphClient.Users[userId].Calendar.CalendarView.GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.StartDateTime = DateTime.Now.AddMonths(-1).ToUniversalTime()
                         .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");
                    requestConfiguration.QueryParameters.EndDateTime = DateTime.Now.AddMonths(6).ToUniversalTime()
                         .ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff'Z'");
                    requestConfiguration.QueryParameters.Filter = "contains(subject,'Urlaub') or contains(subject,'Vacation') or contains(subject,'urlaub') or contains(subject,'vacation')";
                });
                if (userEvents == null)
                    continue;
                var eventList = new List<Event>();
                var eventIterator = PageIterator<Event, EventCollectionResponse>.CreatePageIterator(
                    graphClient, userEvents, (e) =>
                    {
                        if (e.Organizer?.EmailAddress?.Address == userEmail && e.Sensitivity != Sensitivity.Private)
                        {
                            eventList.Add(e);
                        }
                        return true;
                    }
                );
                await eventIterator.IterateAsync();
                events.AddRange(eventList);
            }
            catch (Exception ex)
            {
                HandleException($"Couldn't get events for {userEmail}", ex);
            }
        }
    }

    return events;
}

static GraphServiceClient GetGraphClient(string clientId, string tenantId)
{
    var scopes = new[] { "User.Read", "Calendars.ReadWrite.Shared", "Group.ReadWrite.All" };

    if (Environment.GetEnvironmentVariable("DOTNET_RUNNING_IN_CONTAINER") == "true")
    {
        var options = new DeviceCodeCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            ClientId = clientId,
            TenantId = tenantId,
            DeviceCodeCallback = (code, cancellation) =>
            {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            },
        };

        var deviceCodeCredential = new DeviceCodeCredential(options);

        return new GraphServiceClient(deviceCodeCredential, scopes);
    }
    else
    {
        var options = new InteractiveBrowserCredentialOptions
        {
            TenantId = tenantId,
            ClientId = clientId,
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            RedirectUri = new Uri("http://localhost"),
        };

        var interactiveCredential = new InteractiveBrowserCredential(options);

        return new GraphServiceClient(interactiveCredential, scopes);
    }
}

static async Task<string?> GetGroupId(string groupName, GraphServiceClient graphClient)
{
    string? groupId = null;
    try
    {
        var group = await graphClient.Groups.GetAsync(requestConfiguration =>
            requestConfiguration.QueryParameters.Filter = $"displayName eq '{groupName}'"
        );
        groupId = (group?.Value?.FirstOrDefault()?.Id) ?? throw new Exception($"Could not find group with name {groupName}");
    }
    catch (Exception ex)
    {
        HandleException($"Couldn't get group {groupName}", ex);
    }

    return groupId;
}

static async Task<List<DirectoryObject>> GetMembers(string groupId, GraphServiceClient graphClient)
{
    List<DirectoryObject>? members = null;
    try
    {
        var membersResult = await graphClient.Groups[groupId].Members.GetAsync();
        members = membersResult?.Value;
    }
    catch (Exception ex)
    {
        HandleException($"Couldn't get members for {groupId}", ex);
    }

    return members ?? new List<DirectoryObject>();
}

static async Task CleanCalendar(string groupId, GraphServiceClient graphClient)
{
    var entries = await graphClient.Groups[groupId].Calendar.Events.GetAsync();
    if (entries?.Value == null) return;
    foreach (var entry in entries.Value)
    {
        try
        {
            await graphClient.Groups[groupId].Calendar.Events[entry.Id].DeleteAsync();
        }
        catch (Exception ex)
        {
            HandleException($"Couldn't delete event {entry.Id}", ex);
        }
    }
}

static void HandleException(string msg, Exception ex)
{
    Console.WriteLine($"{msg}: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
    if (ex != null && ex is ODataError)
    {
        var oe = ex as ODataError;
        if (oe!.Error != null)
            Console.WriteLine($"\t{oe.Error.Message}");
    }
}