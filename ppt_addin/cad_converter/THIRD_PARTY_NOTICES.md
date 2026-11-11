# DWG to SVG converter dependencies

The converter uses the following pinned packages.

| Package | Version | License used for this open source distribution |
| --- | --- | --- |
| ACadSharp.Image | 0.1.8 | MIT |
| ACadSharp | 3.8.35 | MIT |
| SixLabors.Fonts | 3.2.0 | Apache License 2.0 under the Six Labors Split License |
| SixLabors.ImageSharp | 4.1.3 | Apache License 2.0 under the Six Labors Split License |
| SixLabors.ImageSharp.Drawing | 3.1.3 | Apache License 2.0 under the Six Labors Split License |
| SixLabors.PolygonClipper | 1.0.2 | Apache License 2.0 under the Six Labors Split License |
| System.IO.Hashing | 8.0.0 | MIT |
| .NET runtime for self-contained Windows builds | 8.0.34 | MIT and upstream third-party notices |

The version-specific Six Labors licenses grant Apache License 2.0 use for software licensed under an open source or source available license. This converter is distributed with the MIT-licensed PowerPoint toolbox and its source. The Six Labors dependencies retain their own copyright and license terms.

Six Labors packages are consumed through ACadSharp.Image's published NuGet dependency graph. Its package metadata excludes dependency Build and Analyzers assets. The converter preserves that standard transitive package relationship, and packages.lock.json pins the resolved packages. No Six Labors license-validation task is edited or disabled and no upstream author's license certificate is copied into this project. The version-specific Split License also explicitly grants Apache License 2.0 use for a transitive package dependency.

Source and license references are available at the following locations.

- https://github.com/slaveOftime/ACadSharp.Image
- https://github.com/DomCR/ACadSharp
- https://github.com/SixLabors/Fonts/blob/v3.2.0/LICENSE
- https://github.com/SixLabors/ImageSharp/blob/v4.1.3/LICENSE
- https://github.com/SixLabors/ImageSharp.Drawing/blob/v3.1.3/LICENSE
- https://github.com/SixLabors/PolygonClipper
- https://www.apache.org/licenses/LICENSE-2.0

Copies of the upstream package license files are included in the packaged converter's licenses directory. ACadSharp.Image and ACadSharp are copyright their respective upstream contributors. SixLabors dependencies are copyright Six Labors. This distribution does not relicense those components under the toolbox's MIT license.
