using Cis.Application.Bordx;
using Cis.Domain.Errors;
using Cis.Infrastructure.Bordx;

namespace Cis.Api.Tests;

/// <summary>Bordx monitoring against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class BordxMonitoringDatabaseTests
{
    [Fact]
    public async Task Control_outstanding_and_worksheets()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BordxMonitoringService(TestDb.Sessions(), new BordxMonitoringRepository(), TimeProvider.System);
        var control = await service.ControlAsync(new BordxControlRequest { Company = "DAYANG ENTERPRISE SDN BHD", DateFrom = "01/01/2021" });
        Assert.Equal(17, control.Rows.Count);
        var os = await service.OutstandingAsync("ALAM TEKNOKRAT SDN BHD", "01/01/2022");
        Assert.Equal(4, os.Rows.Count);
        Assert.Equal(os.Total, os.Total14 + os.TotalOver14);
        Assert.Equal(BordxMonitoringService.NoRecordCompany,
            (await Assert.ThrowsAsync<NotFoundException>(() => service.OutstandingAsync("DAYANG ENTERPRISE SDN BHD", null))).Message);
        var wks = await service.WorksheetsAsync(new BordxWksRequest { Company = "DAYANG ENTERPRISE SDN BHD", BordxFrom = "KDC55049", BordxTo = "KDC55049" });
        Assert.All(wks.Rows, x => Assert.Equal("KDC55049", ((string)x["BordxRefNo"]!).Trim()));
    }
}
