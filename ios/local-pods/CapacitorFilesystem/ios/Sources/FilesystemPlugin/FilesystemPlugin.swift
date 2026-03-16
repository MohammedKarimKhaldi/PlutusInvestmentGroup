import Foundation
import Capacitor

@objc(FilesystemPlugin)
public class FilesystemPlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "FilesystemPlugin"
    public let jsName = "Filesystem"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "writeFile", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "checkPermissions", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "requestPermissions", returnType: CAPPluginReturnPromise)
    ]

    @objc override public func checkPermissions(_ call: CAPPluginCall) {
        call.resolve(["publicStorage": "granted"])
    }

    @objc override public func requestPermissions(_ call: CAPPluginCall) {
        call.resolve(["publicStorage": "granted"])
    }

    @objc func writeFile(_ call: CAPPluginCall) {
        guard let targetPath = call.getString("path"), !targetPath.isEmpty else {
            call.reject("A file path is required.")
            return
        }

        guard let rawData = call.getString("data") else {
            call.reject("File data is required.")
            return
        }

        do {
            let fileURL = try resolveURL(path: targetPath, directory: call.getString("directory"))
            let parentDirectory = fileURL.deletingLastPathComponent()

            if call.getBool("recursive") == true {
                try FileManager.default.createDirectory(at: parentDirectory, withIntermediateDirectories: true)
            }

            let payload = try decodePayload(rawData, encoding: call.getString("encoding"))
            try payload.write(to: fileURL, options: .atomic)

            call.resolve(["uri": fileURL.absoluteString])
        } catch {
            call.reject(error.localizedDescription)
        }
    }

    private func resolveURL(path: String, directory: String?) throws -> URL {
        let cleanPath = path.trimmingCharacters(in: .whitespacesAndNewlines)

        if cleanPath.hasPrefix("file://"), let url = URL(string: cleanPath) {
            return url
        }

        if cleanPath.hasPrefix("/") {
            return URL(fileURLWithPath: cleanPath, isDirectory: false)
        }

        let baseDirectory: FileManager.SearchPathDirectory
        switch directory?.lowercased() {
        case "documents":
            baseDirectory = .documentDirectory
        case "data":
            baseDirectory = .applicationSupportDirectory
        default:
            baseDirectory = .cachesDirectory
        }

        guard let baseURL = FileManager.default.urls(for: baseDirectory, in: .userDomainMask).first else {
            throw NSError(domain: "FilesystemPlugin", code: 1, userInfo: [NSLocalizedDescriptionKey: "Unable to resolve a writable directory."])
        }

        return baseURL.appendingPathComponent(cleanPath, isDirectory: false)
    }

    private func decodePayload(_ value: String, encoding: String?) throws -> Data {
        if let encoding, encoding.lowercased().contains("utf") {
            return Data(value.utf8)
        }

        if let commaIndex = value.firstIndex(of: ","), value[..<commaIndex].contains(";base64") {
            let base64 = String(value[value.index(after: commaIndex)...])
            if let data = Data(base64Encoded: base64, options: .ignoreUnknownCharacters) {
                return data
            }
        }

        if let data = Data(base64Encoded: value, options: .ignoreUnknownCharacters) {
            return data
        }

        return Data(value.utf8)
    }
}
