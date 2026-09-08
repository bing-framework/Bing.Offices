using System;
using System.Linq;
using System.Reflection;
using Bing.Offices.Exceptions;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// SheetExtensions 的图片添加、读取、筛选、移除和移动职责测试。
/// </summary>
public sealed class NpoiSheetPictureExtensionsTest
{
    [Theory]
    [InlineData(false, "png")]
    [InlineData(false, "jpeg")]
    [InlineData(false, "gif")]
    [InlineData(false, "unknown")]
    [InlineData(true, "png")]
    [InlineData(true, "jpeg")]
    [InlineData(true, "gif")]
    [InlineData(true, "unknown")]
    public void AddPicture_ShouldSupportBothProvidersAndResolveSignatures(bool isXlsx, string signature)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");
        var style = new PictureStyle
        {
            AnchorDx1 = 11,
            AnchorDx2 = 22,
            AnchorDy1 = 33,
            AnchorDy2 = 44
        };
        var pictureBytes = GetPictureBytes(signature);

        if (!isXlsx && signature == "gif")
        {
            var unsupported = Assert.Throws<InvalidOperationException>(() =>
                SheetExtensions.AddPicture(sheet, new PictureInfo(2, 4, 3, 5, pictureBytes, style)));

            Assert.Equal("Unexpected picture format: GIF", unsupported.Message);
            Assert.Empty(workbook.GetAllPictures());
            return;
        }

        SheetExtensions.AddPicture(sheet, new PictureInfo(2, 4, 3, 5, pictureBytes, style));

        var pictureData = Assert.IsAssignableFrom<IPictureData>(Assert.Single(workbook.GetAllPictures()));
        var pictureInfo = Assert.Single(SheetExtensions.GetAllPictureInfos(sheet));

        Assert.Equal(GetExpectedPictureType(signature), pictureData.PictureType);
        Assert.Equal(pictureBytes, pictureData.Data);
        Assert.Equal(2, pictureInfo.MinRow);
        Assert.Equal(4, pictureInfo.MaxRow);
        Assert.Equal(3, pictureInfo.MinCol);
        Assert.Equal(5, pictureInfo.MaxCol);
        Assert.Equal(pictureBytes, pictureInfo.PictureData);
        Assert.Equal(11, pictureInfo.PictureStyle.AnchorDx1);
        Assert.Equal(22, pictureInfo.PictureStyle.AnchorDx2);
        Assert.Equal(33, pictureInfo.PictureStyle.AnchorDy1);
        Assert.Equal(44, pictureInfo.PictureStyle.AnchorDy2);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddPicture_ShouldRejectInvalidPictureInfoBeforeWorkbookMutation(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");
        var style = new PictureStyle();

        Assert.Equal("picInfo", Assert.Throws<ArgumentNullException>(() =>
            SheetExtensions.AddPicture(sheet, null)).ParamName);
        Assert.Equal("picInfo", Assert.Throws<ArgumentException>(() => SheetExtensions.AddPicture(sheet,
            new PictureInfo(0, 1, 0, 1, null, style))).ParamName);
        Assert.Equal("picInfo", Assert.Throws<ArgumentException>(() => SheetExtensions.AddPicture(sheet,
            new PictureInfo(0, 1, 0, 1, Array.Empty<byte>(), style))).ParamName);
        Assert.Equal("picInfo", Assert.Throws<ArgumentException>(() => SheetExtensions.AddPicture(sheet,
            new PictureInfo(0, 1, 0, 1, PngBytes(), null))).ParamName);

        Assert.Empty(workbook.GetAllPictures());
        Assert.Empty(SheetExtensions.GetAllPictureInfos(sheet));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GetAllPictureInfos_ShouldFilterByIntersectionAndOnlyInternal(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");

        Assert.Empty(SheetExtensions.GetAllPictureInfos(sheet));
        AddPicture(sheet, 1, 2, 1, 2);
        AddPicture(sheet, 4, 5, 1, 2);

        var internalOnly = SheetExtensions.GetAllPictureInfos(sheet, 1, 2, 1, 2);
        var intersecting = SheetExtensions.GetAllPictureInfos(sheet, 2, 5, 1, 2, false);
        var noMatch = SheetExtensions.GetAllPictureInfos(sheet, 10, 12, 10, 12, false);

        Assert.Single(internalOnly);
        Assert.Equal(1, internalOnly[0].MinRow);
        Assert.Equal(2, intersecting.Count);
        Assert.Contains(intersecting, picture => picture.MinRow == 1 && picture.MaxRow == 2);
        Assert.Contains(intersecting, picture => picture.MinRow == 4 && picture.MaxRow == 5);
        Assert.Empty(noMatch);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovePictures_ShouldSupportLocalIntersectionAndFullRemoval(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");
        AddPicture(sheet, 1, 2, 1, 2);
        AddPicture(sheet, 4, 5, 1, 2);
        AddPicture(sheet, 8, 9, 1, 2);

        SheetExtensions.RemovePictures(sheet, 1, 2, 1, 2);
        Assert.Equal(2, SheetExtensions.GetAllPictureInfos(sheet).Count);

        SheetExtensions.RemovePictures(sheet, 2, 5, 1, 2, false);
        var remaining = SheetExtensions.GetAllPictureInfos(sheet);
        Assert.Single(remaining);
        Assert.Equal(8, remaining[0].MinRow);

        SheetExtensions.RemovePictures(sheet);

        Assert.Empty(SheetExtensions.GetAllPictureInfos(sheet));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MovePictures_ShouldMoveAllAndSelectedAnchors(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");
        AddPicture(sheet, 1, 2, 1, 2);
        AddPicture(sheet, 4, 5, 4, 5);

        SheetExtensions.MovePictures(sheet, 2, 1);

        AssertPictureBounds(SheetExtensions.GetAllPictureInfos(sheet).OrderBy(x => x.MinRow).ToArray(),
            new[] { 3, 4, 2, 3 }, new[] { 6, 7, 5, 6 });

        SheetExtensions.MovePictures(sheet, 3, 4, 2, 3, true, -1, -1);

        AssertPictureBounds(SheetExtensions.GetAllPictureInfos(sheet).OrderBy(x => x.MinRow).ToArray(),
            new[] { 2, 3, 1, 2 }, new[] { 6, 7, 5, 6 });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TryAddPicture_ShouldSupportByteAndPictureDataOverloads(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");

        Assert.True(SheetExtensions.TryAddPicture(sheet, 1, 2, PngBytes()));

        workbook.AddPicture(PngBytes(), PictureType.PNG);
        var pictureData = GetLastPictureData(workbook);
        Assert.True(SheetExtensions.TryAddPicture(sheet, 3, 4, pictureData));

        Assert.Equal(3, workbook.GetAllPictures().Count);
        Assert.Equal(2, SheetExtensions.GetAllPictureInfos(sheet).Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TryAddPicture_ShouldRejectInvalidArgumentsWithoutAddingPicture(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");

        Assert.Equal("pictureData", Assert.Throws<ArgumentNullException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, (IPictureData)null)).ParamName);
        Assert.Equal("pictureBytes", Assert.Throws<ArgumentNullException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, (byte[])null)).ParamName);
        Assert.Equal("pictureBytes", Assert.Throws<ArgumentException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, Array.Empty<byte>())).ParamName);
        Assert.Equal("row", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.TryAddPicture(sheet, -1, 0, PngBytes())).ParamName);
        Assert.Equal("col", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, -1, PngBytes())).ParamName);
        Assert.Equal("pictureType", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, PngBytes(), (PictureType)int.MaxValue)).ParamName);
        Assert.Equal("sheet", Assert.Throws<ArgumentNullException>(() =>
            SheetExtensions.TryAddPicture((ISheet)null, 0, 0, PngBytes())).ParamName);
        Assert.Empty(workbook.GetAllPictures());
        Assert.Empty(SheetExtensions.GetAllPictureInfos(sheet));

        workbook.AddPicture(PngBytes(), PictureType.PNG);
        var pictureData = GetLastPictureData(workbook);
        var countBeforeObjectValidation = workbook.GetAllPictures().Count;
        Assert.Equal("row", Assert.Throws<ArgumentOutOfRangeException>(() =>
            SheetExtensions.TryAddPicture(sheet, -1, 0, pictureData)).ParamName);
        Assert.Equal(countBeforeObjectValidation, workbook.GetAllPictures().Count);
        Assert.Empty(SheetExtensions.GetAllPictureInfos(sheet));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void TryAddPicture_PostMutationFailure_ShouldWrapAndPreserveWorkbookSideEffect(bool isXlsx,
        bool throwArgumentException)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");
        var adapter = new PostMutationFailureAdapter(throwArgumentException);

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, PngBytes(), PictureType.PNG, adapter));

        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Write, exception.Stage);
        Assert.Equal(throwArgumentException ? typeof(ArgumentException) : typeof(InvalidOperationException),
            exception.InnerException.GetType());
        Assert.Single(workbook.GetAllPictures());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TryAddPicture_ResizeFailure_ShouldWrapAndPreserveWorkbookSideEffect(bool isXlsx)
    {
        using var workbook = CreateWorkbook(isXlsx);
        var sheet = workbook.CreateSheet("Pictures");

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            SheetExtensions.TryAddPicture(sheet, 0, 0, PngBytes(), PictureType.PNG,
                new ResizeFailureAdapter()));

        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Write, exception.Stage);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Single(workbook.GetAllPictures());
    }

    [Fact]
    public void PictureExtensions_UnknownSheetAndWorkbook_ShouldKeepCurrentExceptionTypes()
    {
        var unknownSheet = DispatchProxy.Create<ISheet, UnknownSheetProxy>();

        Assert.Throws<NotSupportedException>(() => SheetExtensions.GetAllPictureInfos(unknownSheet));
        Assert.Throws<NotSupportedException>(() => SheetExtensions.RemovePictures(unknownSheet));
        var unsupported = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            SheetExtensions.MovePictures(unknownSheet));
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, unsupported.Code);
        Assert.Equal(BingOfficesOperation.Export, unsupported.Operation);
        Assert.Equal(BingOfficesStage.Write, unsupported.Stage);
        Assert.Equal("sheet", Assert.Throws<ArgumentNullException>(() =>
            SheetExtensions.MovePictures((ISheet)null)).ParamName);

        var unknownWorkbook = DispatchProxy.Create<IWorkbook, UnknownWorkbookProxy>();
        UnknownSheetWithWorkbookProxy.WorkbookValue = unknownWorkbook;
        try
        {
            var workbookBackedSheet = DispatchProxy.Create<ISheet, UnknownSheetWithWorkbookProxy>();
            var picture = new PictureInfo(0, 1, 0, 1, PngBytes(), new PictureStyle());

            Assert.Throws<NullReferenceException>(() =>
                SheetExtensions.AddPicture(workbookBackedSheet, picture));
            Assert.Throws<NullReferenceException>(() =>
                SheetExtensions.TryAddPicture(workbookBackedSheet, 0, 0, PngBytes()));
        }
        finally
        {
            UnknownSheetWithWorkbookProxy.WorkbookValue = null;
        }
    }

    private static void AddPicture(ISheet sheet, int minRow, int maxRow, int minCol, int maxCol)
    {
        SheetExtensions.AddPicture(sheet, new PictureInfo(minRow, maxRow, minCol, maxCol, PngBytes(),
            new PictureStyle()));
    }

    private static IWorkbook CreateWorkbook(bool isXlsx) =>
        isXlsx ? new XSSFWorkbook() : new HSSFWorkbook();

    private static byte[] PngBytes() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    private static byte[] GifBytes() => Convert.FromBase64String(
        "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

    private static byte[] GetPictureBytes(string signature) => signature switch
    {
        "png" => PngBytes(),
        "jpeg" => new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
        "gif" => GifBytes(),
        "unknown" => new byte[] { 0x10, 0x20, 0x30, 0x40 },
        _ => throw new ArgumentOutOfRangeException(nameof(signature))
    };

    private static PictureType GetExpectedPictureType(string signature) => signature switch
    {
        "png" => PictureType.PNG,
        "jpeg" => PictureType.JPEG,
        "gif" => PictureType.GIF,
        "unknown" => PictureType.PNG,
        _ => throw new ArgumentOutOfRangeException(nameof(signature))
    };

    private static IPictureData GetLastPictureData(IWorkbook workbook)
    {
        var pictures = workbook.GetAllPictures();
        return Assert.IsAssignableFrom<IPictureData>(pictures[pictures.Count - 1]);
    }

    private static void AssertPictureBounds(PictureInfo[] actual, params int[][] expected)
    {
        Assert.Equal(expected.Length, actual.Length);
        for (var i = 0; i < expected.Length; i++)
        {
            Assert.Equal(expected[i][0], actual[i].MinRow);
            Assert.Equal(expected[i][1], actual[i].MaxRow);
            Assert.Equal(expected[i][2], actual[i].MinCol);
            Assert.Equal(expected[i][3], actual[i].MaxCol);
        }
    }

    private sealed class PostMutationFailureAdapter : IPictureMutationAdapter
    {
        private readonly bool _throwArgumentException;

        public PostMutationFailureAdapter(bool throwArgumentException)
        {
            _throwArgumentException = throwArgumentException;
        }

        public int AddPicture(ISheet sheet, byte[] pictureBytes, PictureType pictureType) =>
            sheet.Workbook.AddPicture(pictureBytes, pictureType);

        public IClientAnchor CreateClientAnchor(ISheet sheet) =>
            sheet.Workbook.GetCreationHelper().CreateClientAnchor();

        public IDrawing GetOrCreateDrawing(ISheet sheet) =>
            sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();

        public IPicture CreatePicture(IDrawing drawing, IClientAnchor anchor, int pictureIndex) =>
            throw (_throwArgumentException
                ? new ArgumentException("测试形状创建失败")
                : new InvalidOperationException("测试形状创建失败"));

        public void Resize(IPicture picture)
        {
        }
    }

    private sealed class ResizeFailureAdapter : IPictureMutationAdapter
    {
        public int AddPicture(ISheet sheet, byte[] pictureBytes, PictureType pictureType) =>
            sheet.Workbook.AddPicture(pictureBytes, pictureType);

        public IClientAnchor CreateClientAnchor(ISheet sheet) =>
            sheet.Workbook.GetCreationHelper().CreateClientAnchor();

        public IDrawing GetOrCreateDrawing(ISheet sheet) =>
            sheet.DrawingPatriarch ?? sheet.CreateDrawingPatriarch();

        public IPicture CreatePicture(IDrawing drawing, IClientAnchor anchor, int pictureIndex) =>
            drawing.CreatePicture(anchor, pictureIndex);

        public void Resize(IPicture picture) =>
            throw new InvalidOperationException("测试图片调整大小失败");
    }

    private class UnknownSheetProxy : DispatchProxy
    {
        protected override object Invoke(MethodInfo targetMethod, object[] args) =>
            targetMethod.ReturnType.IsValueType ? Activator.CreateInstance(targetMethod.ReturnType) : null;
    }

    private class UnknownWorkbookProxy : DispatchProxy
    {
        protected override object Invoke(MethodInfo targetMethod, object[] args) =>
            targetMethod.ReturnType.IsValueType ? Activator.CreateInstance(targetMethod.ReturnType) : null;
    }

    private class UnknownSheetWithWorkbookProxy : DispatchProxy
    {
        public static IWorkbook WorkbookValue { get; set; }

        protected override object Invoke(MethodInfo targetMethod, object[] args) =>
            targetMethod.Name == "get_Workbook"
                ? WorkbookValue
                : targetMethod.ReturnType.IsValueType
                    ? Activator.CreateInstance(targetMethod.ReturnType)
                    : null;
    }
}
