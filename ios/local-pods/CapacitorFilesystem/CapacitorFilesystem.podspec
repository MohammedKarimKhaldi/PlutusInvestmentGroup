Pod::Spec.new do |s|
  s.name = 'CapacitorFilesystem'
  s.version = '8.1.2'
  s.summary = 'Local lightweight Filesystem plugin for iOS and Mac Catalyst.'
  s.license = 'MIT'
  s.homepage = 'https://capacitorjs.com'
  s.author = 'Plutus Investment Group'
  s.source = { :path => '.' }
  s.source_files = 'ios/Sources/**/*.{swift,h,m}'
  s.ios.deployment_target = '15.0'
  s.dependency 'Capacitor'
  s.swift_version = '5.1'
end
