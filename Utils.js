class Utils {
  static getFileIdByUrl(url) {
    // https://docs.google.com/spreadsheets/d/{SheetId}/edit#gid={gid}
    return url.split('/')[5];
  }
}
