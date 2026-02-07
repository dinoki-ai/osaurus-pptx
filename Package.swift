// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "osaurus-pptx",
    platforms: [.macOS(.v15)],
    products: [
        .library(name: "osaurus-pptx", type: .dynamic, targets: ["osaurus_pptx"])
    ],
    targets: [
        .target(
            name: "osaurus_pptx",
            path: "Sources/osaurus_pptx"
        ),
        .testTarget(
            name: "osaurus_pptx_tests",
            dependencies: ["osaurus_pptx"],
            path: "Tests/osaurus_pptx_tests"
        )
    ]
)