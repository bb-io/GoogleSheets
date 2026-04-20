using Apps.GoogleSheets.Models.Dto;

namespace Apps.GoogleSheets.Models.Responses;

public record SheetNamesResponse(List<SheetNameDto> Sheets);