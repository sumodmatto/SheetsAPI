// src/Spreadsheet.ts
export class Spreadsheet {
  private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet

  constructor(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
    this.spreadsheet = spreadsheet
  }

  /**
   * Declare addDeveloperMetadata method signature
   * @param key
   */
  addDeveloperMetadata(key: string): Spreadsheet
  addDeveloperMetadata(
    key: string,
    visibility: GoogleAppsScript.Spreadsheet.DeveloperMetadataVisibility
  ): Spreadsheet
  addDeveloperMetadata(key: string, value: string): Spreadsheet
  addDeveloperMetadata(
    key: string,
    value: string,
    visibility: GoogleAppsScript.Spreadsheet.DeveloperMetadataVisibility
  ): Spreadsheet

  /**
   * Add developer metadata to the spreadsheet.
   * @param key
   * @param valueOrVisibility
   * @param visibility
   * @returns
   */
  addDeveloperMetadata(
    key: string,
    valueOrVisibility?: string | GoogleAppsScript.Spreadsheet.DeveloperMetadataVisibility,
    visibility?: GoogleAppsScript.Spreadsheet.DeveloperMetadataVisibility
  ): Spreadsheet {
    if (typeof valueOrVisibility === 'string' && visibility !== undefined) {
      // Handle as value + visibility (assuming visibility is enum and valid)
      this.spreadsheet.addDeveloperMetadata(key, valueOrVisibility, visibility)
    } else if (typeof valueOrVisibility === 'string') {
      // Handle as value only
      this.spreadsheet.addDeveloperMetadata(key, valueOrVisibility)
    } else if (
      typeof valueOrVisibility === 'number' &&
      Object.values(GoogleAppsScript.Spreadsheet.DeveloperMetadataVisibility).includes(
        valueOrVisibility
      )
    ) {
      // Handle as visibility only, assuming it's a valid enum value
      this.spreadsheet.addDeveloperMetadata(key, valueOrVisibility)
    } else {
      // Default action
      this.spreadsheet.addDeveloperMetadata(key)
    }
    return this
  }

  /**
   * Declare addEditor method signature
   */
  addEditor(emailAddress: string): Spreadsheet
  addEditor(user: GoogleAppsScript.Base.User): Spreadsheet

  /**
   * Add an editor to the spreadsheet.
   * @param emailAddressOrUser
   * @returns
   */
  addEditor(emailAddressOrUser: string | GoogleAppsScript.Base.User): Spreadsheet {
    if (typeof emailAddressOrUser === 'string') {
      this.spreadsheet.addEditor(emailAddressOrUser)
    } else {
      this.spreadsheet.addEditor(emailAddressOrUser)
    }
    return this
  }

  /**
   * Add editors to the spreadsheet by email address.
   * @returns Spreadsheet
   */
  addEditors(emailAddresses: string[]): Spreadsheet {
    this.spreadsheet.addEditors(emailAddresses)
    return this
  }

  /**
   * Add a menu to the spreadsheet.
   * @param name
   * @param subMenus
   * @returns Spreadsheet
   */
  addMenu(name: string, subMenus: { name: string; functionName: string }[]): Spreadsheet {
    this.spreadsheet.addMenu(name, subMenus)
    return this
  }

  /**
   * Declare addViewer method signature
   */
  addViewer(emailAddress: string): Spreadsheet
  addViewer(user: GoogleAppsScript.Base.User): Spreadsheet

  /**
   * Add a viewer to the spreadsheet.
   * @param emailAddressOrUser
   * @returns
   */
  addViewer(emailAddressOrUser: string | GoogleAppsScript.Base.User): Spreadsheet {
    if (typeof emailAddressOrUser === 'string') {
      this.spreadsheet.addViewer(emailAddressOrUser)
    } else {
      this.spreadsheet.addViewer(emailAddressOrUser)
    }
    return this
  }

  /**
   * Add viewers to the spreadsheet by email address.
   * @returns Spreadsheet
   */
  addViewers(emailAddresses: string[]): Spreadsheet {
    this.spreadsheet.addViewers(emailAddresses)
    return this
  }

  getDeveloperMetadata(): GoogleAppsScript.Spreadsheet.DeveloperMetadata[] {
    return this.spreadsheet.getDeveloperMetadata()
  }

  getName(): string {
    return this.spreadsheet.getName()
  }

  getSheets(): GoogleAppsScript.Spreadsheet.Sheet[] {
    return this.spreadsheet.getSheets()
  }

  getRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range {
    return this.spreadsheet.getRange(a1Notation)
  }

  customMethod(): string {
    return 'This is a custom method.'
  }
}
