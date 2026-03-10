// swift-tools-version:5.3
import PackageDescription

let package = Package(
    name: "capacitor-swift-pm",
    products: [
        .library(
            name: "Capacitor",
            targets: ["Capacitor"]
        ),
        .library(
            name: "Cordova",
            targets: ["Cordova"]
        )
    ],
    dependencies: [],
    targets: [
        .binaryTarget(
            name: "Capacitor",
            path: "Artifacts/Capacitor.xcframework.zip"
        ),
        .binaryTarget(
            name: "Cordova",
            path: "Artifacts/Cordova.xcframework.zip"
        )
    ]
)
