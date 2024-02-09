# PowerPoint `PresentationCloseFinal` event

> The Microsoft PowerPoint on macOS does not fire the `PresentationCloseFinal` event.
> This event should be fired when the presentation document is really closed.
>
> The event is fired on PowerPoint on Windows.

## Installing the macro file

Use Microsoft PowerPoint for Mac, choose **Tools** > **PowerPoint Add-ins...** dialog
to add the compiled `PresentationCloseFinalAddin.ppam` file to the application.


## Bug report

The macro file is an addin which can be added to the Microsoft PowerPoint app
on Windows or on Mac.

The macro sets handlers for various events in the presentation lifecycle:

* `PresentationBeforeClose`
* `PresentationClose`
* `PresentationCloseFinal`

When user closes the presentation document, PowerPoint will fire the `PresentationClose`
event. When document is unsaved, PowerPoint will display a dialog box to save the file
after this event handler code was executed.

If user choses to cancel the save file dialog box, the presentation will still be opened
and user can work with it. This makes the `PresentationClose` unsuitable to clean up
resources.

The PowerPoint 2010 on Windows added new event `PresentationCloseFinal` which is called
if the presentation document is really closed.

It is the right event to clean up any unmaganed resources the macro or add-in may be
working with.


### Expected behavior

The `PresentationCloseFinal` should be fired on PowerPoint on macOS at the same
point when Windows version of PowerPoint does it.

### Actual behavior

The `PresentationCloseFinal` is never fired on PowerPoint on macOS.


## Compiling the source code

1. Install [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download)
2. Install [vbamc](https://www.nuget.org/packages/vbamc) tool
3. Compile addin with `make`

```shell
brew install make
brew install --cask dotnet-sdk
dotnet tool install --global vbamc

cd src
make
```

The compiled `PresentationCloseFinalAddin.ppam` file will be in the `src/bin` directory.


## References

> You probably want PresentationBeforeClose or PresentationCloseFinal which was added in PowerPoint 2010.
>
> <https://stackoverflow.com/questions/9097230/powerpoint-after-presentation-close-event>
