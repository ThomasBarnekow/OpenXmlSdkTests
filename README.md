# Open XML SDK Tests

At the moment, this repository contains projects demonstrating how the Flat OPC and cloning features
evolved over the past couple of releases and how that breaks some code that depends on Flat OPC documents.

As a side effect, it also demonstrates the use of the `GetPackage()` extension method that replaced the
`Package` property of the `OpenXmlPackage` class. The latter was made obsolete in v2.20.0 by way of
creating an error (not a warning, mind you) while not providing any replacement just yet (to my knowledge),
which made v2.20.0 entirely useless for those who needed access to `PackagePart` (now `IPackagePart`)
instances, for example.

## Why is this important?

When developing web-based Office add-ins (i.e., using the "new" JavaScript-based model), what add-ins
will get when asking for an XML representation of a document or parts thereof are Flat OPC documents.
Should developers then want to process those documents using the Open XML SDK, they'll need to be able
to turn those Flat OPC documents into `WordprocessingDocument` instances. This is where the Flat OPC
feature of the Open XML SDK comes in.

Next, in their processing pipelines, developers might want to "clone" those documents. This allows them
to create a copy in memory or write that document to a stream (e.g., `MemoryStream`, `FileStream`),
for example.

## What has changed?

Here's a short history from my point of view:

- In **v2.19.0**, everything works as expected.
- In **v2.20.0**, the `Package` property was made obsolete (which is why I could not use it)
- In **v3.0.0**, after making the necessary changes related to `Package` being obsolete, everything
  works as expected.
- In **v3.0.1**, what looks like a bug was introduced in the Flat OPC feature. The package was
  suddenly missing the relationship packages. For some reason, this also breaks the cloning feature,
  so you can no longer copy or save `WordprocessingDocument` instances that began their lives as Flat
  OPC documents.
- In **v3.1.1**, the latest version as of October 27, 2024, the same problem persists.
