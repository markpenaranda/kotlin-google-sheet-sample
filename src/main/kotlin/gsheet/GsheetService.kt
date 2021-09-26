package gsheet


import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.JsonFactory
import com.google.api.client.json.jackson2.JacksonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.services.drive.Drive
import com.google.api.services.drive.DriveScopes
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.api.services.sheets.v4.model.Spreadsheet
import com.google.api.services.sheets.v4.model.SpreadsheetProperties
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.InputStreamReader
import java.io.OutputStream
import java.util.UUID


class GsheetService {

    private lateinit var sheetsService: Sheets

    fun auth() {
        sheetsService = Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
            .setApplicationName(APPLICATION_NAME)
            .build()
    }

    private fun <T> driveAuth(action: (Drive) -> T): T {
      return Drive.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT, DriveScopes.all()))
            .setApplicationName(APPLICATION_NAME)
            .build()
            .let { action(it) }
    }

    fun createWorkbook(title: String) {
        if(!this::sheetsService.isInitialized) {
            throw IllegalAccessError("No Auth")
        }

        Spreadsheet()
            .setProperties(
                SpreadsheetProperties()
                    .setTitle(title)
            ).let {
               sheetsService.spreadsheets().create(it)
                   .setFields("spreadsheetId")
                   .execute();
           }
    }

    fun getSpreadsheetDetails(spreadsheetId: String): Workbook {
       return driveAuth {
           it.Files().get(spreadsheetId).setFields("name,size,modifiedTime").execute().let { file ->
                Workbook(
                    file.name,
                    file.size.toLong(),
                    file.modifiedTime.value
                )
            }
        }

    }

    fun getLatestWorkbook(spreadsheetId: String): Filename {
        val fileName = Filename.generate()
        val outputStream: OutputStream = FileOutputStream(fileName.value)
        driveAuth {
            it.Files()
                .export(spreadsheetId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .executeMediaAndDownloadTo(outputStream)
        }

        return fileName
    }

    fun updateFile() {

    }


    private fun getCredentials(HTTP_TRANSPORT: NetHttpTransport, scope: Set<String> = SCOPES): Credential {
        // Load client secrets.
        val credentials = GsheetService.javaClass.getResourceAsStream(CREDENTIALS_FILE_PATH)
            ?: throw FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH)
        val clientSecrets: GoogleClientSecrets = GoogleClientSecrets.load(JSON_FACTORY, InputStreamReader(credentials))

        // Build flow and trigger user authorization request.
        val flow: GoogleAuthorizationCodeFlow = GoogleAuthorizationCodeFlow.Builder(
            HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES
        )
            .setDataStoreFactory(FileDataStoreFactory(File(TOKENS_DIRECTORY_PATH)))
            .setAccessType("offline")
            .build()
        val receiver: LocalServerReceiver = LocalServerReceiver.Builder().setPort(8888).build()
        return AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
    }

    companion object {
        private val SCOPES: Set<String> = SheetsScopes.all()
        private const val CREDENTIALS_FILE_PATH = "/credentials.json"
        private val JSON_FACTORY: JsonFactory = JacksonFactory.getDefaultInstance()
        private const val TOKENS_DIRECTORY_PATH = "tokens"
        val HTTP_TRANSPORT: NetHttpTransport = GoogleNetHttpTransport.newTrustedTransport()
        const val APPLICATION_NAME = "Test App"
    }

}

data class Filename(val value: String) {
    companion object {
        fun generate(): Filename {
           return Filename(UUID.randomUUID().toString() + ".xlsx")
        }
    }
}

data class Workbook(
    val title: String,
    val fileSize: Long,
    val updatedAt: Long
)
