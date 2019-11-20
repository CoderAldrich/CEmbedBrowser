/* Events.c
 
This is a Win32 C application (ie, no MFC, WTL, nor even any C++ -- just plain C) that demonstrates
how to embed a browser "control" (actually, an OLE object) in your own window (in order to display a
web page, or an HTML file on disk).

This example displays a HTML string (in memory) in a browser object. This string contains three
elements.

The first element is a FONT tag. It simply colors the text "This is a test of events"
red. We tell the browser that we want it to inform us when the user moves the mouse over that
text (which is encompassed by the FONT tag). To do this we need to "connect" to that FONT tag's
"mouseover" event. Whenever the user moves the mouse over whatever is inside of that FONT tag,
then the browser will call one of our functions. Actually, it will call a function inside of
CWebPage.DLL (on our behalf) which then passes our own window (hosting the browser object) a WM_NOTIFY
message with an LPARAM that is a pointer to a WEBPARAMS struct (defined in CWebPage.h). By looking at
the contents of this struct, we'll know that we've just been informed of the mouse moving over
that FONT tag. Furthermore, the WPARAM of the WM_NOTIFY will be a pointer to an IDispatch object
we can use to call functions that modify this particular FONT tag. At this point, we'll change the color
of the FONT tag (and therefore the color of the text it displays) to "maroon".

We'll also tell the browser that we want it to inform us when the user moves the mouse away from
that same FONT tag). To do this we need to "connect" to that FONT tag's "mouseout" event. Whenever
the user moves the mouse over whatever is inside of that FONT tag (we'll receive a mouseover for
that), and then moves the mouse back outside of that FONT tag, then our own window will be passed
a WM_NOTIFY message with an LPARAM that is a pointer to a WEBPARAMS struct, and a WPARAM that is
an IDispatch object for this particular FONT. By looking at the contents of the WBPARAMS, we'll know
that we've just been informed of the mouse moving back outside that FONT tag. We'll use the
IDispatch to change the color of the FONT tag to seagreen.

The net result is that "This is a test of events" will initially be displayed as red. When the
user moves the mouse over the text, its color will change to maroon, and when the user moves the
mouse off of that text, its color will change to seagreen.

NOTE: In order to be able to connect to events in the FONT tag, we need to give it an "ID" (name).
You do this using the "id" attribute. So here is what our FONT tag looks like in our HTML page:

<FONT id=testfont color=red>This is a test of events</FONT>

Notice that we've given this FONT tag an ID name of "testfont". You can use any name you wish.

For the purposes of CWebPage.DLL informing us of the "mouseover" or "mouseout" event, we also
need to assign a unique, non-zero number to this FONT tag. We'll arbitrary choose 1. We make this
assignment when we call CWebPage.DLL's CreateWebEvtHandler() to connect to this FONT's "mouseover"
and "mouseout" events.

===================================

Our second element is a FORM tag to get some user input. It will have an EDIT control for him to
type in some text and a "submit" button. We tell the browser that we want it to inform us when the
user clicks on the submit button.. To do this we need to "connect" to that FORM tag's "submit" event.
Whenever the user clicks that FORM's submit button, then CWebPage.DLL will pass our own window
(hosting the browser object) a WM_NOTIFY message with an LPARAM that is a pointer to a WEBPARAMS
struct, and a WPARAM that is a pointer to an IDispatch object for this particular FORM. By looking
at the WEBPARAMS, we'll know that we've just been informed of the user clicking on that FORM's submit
button. At this point, we'll use the IDispatch to cancel the submission.

Again, to connect to events in this FORM tag, we need to give it a unique ID (different from the
FONT tag). So here is what our FORM tag looks like in our HTML page:

<FORM id=testform action="http://www.google.com/search">
<INPUT type=text name=testinput>
<INPUT type=submit>
</FORM>
  
Notice that we've given the FORM an ID of "testform".

For the purposes of CWebPage.DLL informing us of the "submit", we also need to assign a unique,
non-zero number to this FORM tag. We'll arbitrary choose 2. We make this assignment when we call
CWebPage.DLL's CreateWebEvtHandler() to connect to this FORM's "submit" event.

===================================

Our third tag is a link the user can click upon to go to some web page. The purpose of this is
to demonstrate how we handle removing our event "connections". After all, the page with our FONT
and FORM will be disappearing as he surfs to another page. So, we want to disconnect from those
elements. When the page changes, CWebPage.DLL will send our own window (hosting the browser) a
WM_NOTIFY event for each element we've connected to. So, we'll receive a WM_NOTIFY for the FONT
element, and another one for the FORM element. Again, we are passed a WEBPARAMS as well as an
IDispatch associated with the element. Unlike with the WEBPARAMS above, this time the struct's "code"
field will be 0. This is an indication that we need to disconnect from the element. We'll use the
IDispatch to call a function to disconnect from whatever event(s) we connected to.

When we get passed a WM_NOTIFY message for each element, telling us to disconnect, what we're
really doing is handling the "beforeunload" event for the web page itself. But you don't need
to know that detail unless you're looking over the CWebPage DLL source.

===================================

There are two more events we connect to. Neither one of these events is associated with a
particular element (tag). Rather, they are applicable to the entire page. We want the browser
to inform us whenever the user double-clicks on the web page. To do this, we connect to the
"dblclick" event.

We also want the browser to inform us when the user uses the scroll bar. To do this, we
connect to the "scroll" event.

For the purposes of CWebPage.DLL informing us of each of these additional two events, we need to
assign a unique, non-zero number to each. We'll arbitrarily choose 3 for the "dblclick" and
4 for the "scroll" event. We make this assignment when we call CWebPage.DLL's CreateWebEvtHandler()
to connect to each of this "global" events.

Whenever the user doubleclicks somewhere on the page, our own window will be sent a WM_NOTIFY with
a WEBPARAMS and an IDispatch associated with the entire web page. By looking at the WEBPARAMS,
we'll be able to tell that the user double-clicked. We'll simply display a message box stating so.

Whenever the user uses the scroll bar, our own window will be sent a WM_NOTIFY with a WEBPARAMS
and an IDispatch associated with the entire web page. By looking at the WEBPARAMS, we'll be able
to tell that the user used the scroll bar. We'll simply display a message box stating so.

And like with the FONT and FORM elements, when the user clicks to move on to another page, we'll
be sent a WM_NOTIFY for each of these events, informing us to disconnect.

===================================

NOTE: If you close the window hosting the web brower, that window will also receive the WM_NOTIFY
messages to disconnect. So this is how consistent cleanup is done. The window is always passed
a WM_NOTIFY for each element event, and/or global event, to which you have connected, whether due
to the window closing, or the page changing.

===================================

This requires IE 5.0 (or better) -- due to how we connect to browser events.
*/

#ifdef UNICODE
#define _UNICODE
#endif
#include <windows.h>
#include "..\cwebpage.h"
#include <tchar.h>



// Handle of our Main Window.
HWND				MainWindow;

// Handle to "CWebPage.DLL"
static HINSTANCE	CWebDll;

// Where we store the pointers to CWebPage.dll's functions
EmbedBrowserObjectPtr		*LpEmbedBrowserObject;
UnEmbedBrowserObjectPtr		*LpUnEmbedBrowserObject;
GetWebPtrsPtr				*LpGetWebPtrs;
DisplayHTMLStrPtr			*LpDisplayHTMLStr;
ResizeBrowserPtr			*LpResizeBrowser;
WaitOnReadyStatePtr			*LpWaitOnReadyState;
GetWebElementPtr			*LpGetWebElement;
CreateWebEvtHandlerPtr		*LpCreateWebEvtHandler;
SetWebReturnValuePtr		*LpSetWebReturnValue;

// The class name of our Main Window. It can be anything of your choosing.
static const TCHAR	ClassName[] = _T("Browser Example");

// Standard Button class name
static const TCHAR	Button[] = _T("Button");

// The class name of our Window to host the browser. It can be anything of your choosing.
static const TCHAR	BrowserClassName[] = _T("Browser Object");
















/****************************** WindowProc() ***************************
 * Our message handler for our Main window.
 */

LRESULT CALLBACK windowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
		case WM_SIZE:
		{
			// Size our browser window to fill the client area of this window.
			// NOTE: We created our child window hosting the browser with an ID of 1000.

			register HWND child = GetDlgItem(hwnd, 1000);
			SetWindowPos(child, 0, 0, 0, LOWORD(lParam), HIWORD(lParam), SWP_NOMOVE|SWP_NOZORDER);
			break;
		}

		case WM_CLOSE:
		{
			// Close this window. This will also cause the child window hosting the browser
			// control to receive its WM_DESTROY
			DestroyWindow(hwnd);

			return(1);
		}

		case WM_DESTROY:
		{
 			// Post the WM_QUIT message to quit the message loop in WinMain()
			PostQuitMessage(0);

			return(1);
		}
	}

	return(DefWindowProc(hwnd, uMsg, wParam, lParam));
}







/************************** browserWindowProc() *************************
 * Our message handler for our window to host the browser.
 */

LRESULT CALLBACK browserWindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
		case WM_NOTIFY:
		{
			register NMHDR	*nmhdr;

			// Is it a message that was sent from one of our _IDispatchEx's (informing us
			// of an event that happened which we want to be informed of)? If so, then the
			// NMHDR->hwndFrom is the handle to our window, and NMHDR->idFrom is 0.
			nmhdr = (NMHDR *)lParam;
			if (((NMHDR *)lParam)->hwndFrom == hwnd && !((NMHDR *)lParam)->idFrom)
			{
				// This message is sent from one of our _IDispatchEx's. The NMHDR is really
				// a WEBPARAMS struct, so we can recast it as so. Also, the wParam is the
				// __IDispatchEx object we used to attach to the event that just happened.
				WEBPARAMS		*webParams;
				_IDispatchEx	*lpDispatch;

				webParams =		(WEBPARAMS *)lParam;
				lpDispatch = 	(_IDispatchEx *)wParam;

				// If NMHDR->code is not zero, then this is not the "beforeunload" event.
				if (((NMHDR *)lParam)->code)
				{
					LPCTSTR		eventType;

					// It is some other event type, such as "onmouseover".
					eventType = webParams->eventStr;

					// Remember that we assigned a unique ID to each element on the page (and
					// each event that is not associated with any particular element). So let's
					// see which element's _IDispatchEx this is.
					switch (lpDispatch->id)
					{
						// Was it our FONT tag?
						case 1:
						{
							// It was the FONT tag. Remember that we told CreateWebEvtHandler() to
							// store this font's IHTMLElement object into our _IDispatchEx's "object"
							// field.
							IHTMLElement *		htmlElem;
							IHTMLFontElement *	htmlFont;
							VARIANT				varColor;

							// It was the FONT tag. Depending upon whether the user is moving the mouse
							// over the font tag, or away from the font tag, we're going to change the
							// font tag's color. (Hey, this is just a demonstration).
							//
							// To change a FONT tag's color, we need to get the FONT element's
							// IHTMLFontElement object so we can call its put_color().
							varColor.vt = VT_BSTR;
							htmlElem = (IHTMLElement *)lpDispatch->object;				

			
							if (htmlElem->lpVtbl->QueryInterface(htmlElem, &IID_IHTMLFontElement, (void **)&htmlFont))
								break;

							// Was it a "mouseover" event?
							if (lstrcmp(eventType, _T("mouseover")) == 0)
							{
								// Change the font tag's color to "maroon". NOTE: We must use
								// a UNICODE string
								varColor.bstrVal = SysAllocString(L"maroon");
							}

							// Was it a "mouseout" event?
							else if (lstrcmp(eventType, _T("mouseout")) == 0)
							{
								// Change the font tag's color to "seagreen"
								varColor.bstrVal = SysAllocString(L"seagreen");
							}

							htmlFont->lpVtbl->put_color(htmlFont, varColor);
							VariantClear(&varColor);

							// Release that IHTMLFontElement
							htmlFont->lpVtbl->Release(htmlFont);

							break;
						}

						// Was it the FORM tag "submit" event?
						case 2:
						{
							// Yes. Let's cancel the submission. To do that, we return a FALSE to the
							// IHTMLFormElement (which was passed to us in the WEBPARAMS->htmlEvent)
							LpSetWebReturnValue((IHTMLEventObj *)webParams->htmlEvent, FALSE);
							break;
						}

						// Was it a scroll event?
						case 3:
						{
							IHTMLWindow2	*htmlWin2;

							htmlWin2 = (IHTMLWindow2 *)lpDispatch->object;				

							// Here, you typically would use the IHTMLWindow2 to call some
							// browser object function to do something with it.

							if (lstrcmp(eventType, _T("scroll")) == 0)
							{
								MessageBox(hwnd, _T("User scrolled"), _T("Event"), 0);
							}

							break;

						}

						// Was it a double-click?
						case 4:
						{
							IHTMLDocument2	*htmlDoc2;

							htmlDoc2 = (IHTMLDocument2 *)lpDispatch->object;				

							// Here, you typically would use the IHTMLDocument2 to call some
							// browser object function to do something with it.
			
							MessageBox(hwnd, _T("User double clicked"), _T("Event"), 0);

							break;
						}
					}
				}

				else
				{
					// This _IDispatch is about to be freed, so we need to detach all
					// events that were attached to it.
					VARIANT			varNull;

					varNull.vt = VT_NULL;

					switch (lpDispatch->id)
					{
						case 1:
						{
							IHTMLElement	*elem;

							elem = (IHTMLElement *)lpDispatch->object;

							// Detach from the "onmouseover" event for this FONT element (by passing
							// put_onmouseover an empty VARIANT)
							elem->lpVtbl->put_onmouseover(elem, varNull);

							// Detach from the "onmouseout" event for this FONT element
							elem->lpVtbl->put_onmouseout(elem, varNull);

							// Note: CWebPage.DLL will release the IHTMLElement for us
		//					elem->lpVtbl->Release(elem);
							break;
						}

						case 2:
						{
							IHTMLFormElement * htmlForm;

							htmlForm = (IHTMLFormElement *)lpDispatch->object;

							// Detach from the "onsubmit" event for this FORM element
							htmlForm->lpVtbl->put_onsubmit(htmlForm, varNull);

		//					htmlForm->lpVtbl->Release(htmlForm);
							break;
						}

						case 3:
						{
							IHTMLWindow2 * htmlWin2;

							htmlWin2 = (IHTMLWindow2 *)lpDispatch->object;

							// Detach from the "onscroll" event for the browser window
							htmlWin2->lpVtbl->put_onscroll(htmlWin2, varNull);

		//					htmlWin2->lpVtbl->Release(htmlWin2);
							break;
						}

						case 4:
						{
							IHTMLDocument2 * htmlDoc2;

							htmlDoc2 = (IHTMLDocument2 *)lpDispatch->object;

							// Detach from the "put_ondblclick" event for the document
							htmlDoc2->lpVtbl->put_ondblclick(htmlDoc2, varNull);

		//					htmlDoc2->lpVtbl->Release(htmlDoc2);
							break;
						}
					}
				}
			}

			// Must be some other entity that sent me a WM_NOTIFY. It wasn't because of a
			// web page action.
			else
			{

			}

			break;
		}

		case WM_SIZE:
		{
			// Size browser object the same size as its host window
			LpResizeBrowser(hwnd, LOWORD(lParam), HIWORD(lParam));
			return(0);
		}

		case WM_CREATE:
		{
			// Embed the browser object into our host window. We need do this only
			// once. Note that the browser object will start calling some of our
			// IOleInPlaceFrame and IOleClientSite functions as soon as we start
			// calling browser object functions in EmbedBrowserObject().
			if (LpEmbedBrowserObject(hwnd)) return(-1);

			// Success
			return(0);
		}

		case WM_DESTROY:
		{
			// Detach the browser object from this window, and free resources.
			LpUnEmbedBrowserObject(hwnd);

			return(TRUE);
		}
	}

	return(DefWindowProc(hwnd, uMsg, wParam, lParam));
}





/****************************** open_webdll() ***************************
 * Opens the CWebPage DLL and gets pointers to functions that we'll be
 * calling.
 *
 * RETURNS: Handle to CWebPage DLL if success, or 0 if an error.
 */

static HINSTANCE open_webdll(void)
{
	// Load our DLL containing the OLE/COM code. We do this once-only. It's named "cwebpage.dll"
	if ((CWebDll = (HINSTANCE)LoadLibrary(_T("cwebpage.dll"))))
	{
		// Get pointers to various functions, and store them in some globals.

		// Get the address of the EmbedBrowserObject() function. NOTE: Only Reginald has this one
		LpEmbedBrowserObject = (EmbedBrowserObjectPtr *)GetProcAddress(CWebDll, EMBEDBROWSEROBJECTNAME);

		// Get the address of the UnEmbedBrowserObject() function. NOTE: Only Reginald has this one
		LpUnEmbedBrowserObject = (UnEmbedBrowserObjectPtr *)GetProcAddress(CWebDll, UNEMBEDBROWSEROBJECTNAME);

		// Get the address of the DisplayHTMLStr() function
		LpDisplayHTMLStr = (DisplayHTMLStrPtr *)GetProcAddress(CWebDll, DISPLAYHTMLSTRNAME);

		// Get the address of the ResizeBrowser() function
		LpResizeBrowser = (ResizeBrowserPtr *)GetProcAddress(CWebDll, RESIZEBROWSERNAME);

		// Get the address of the GetWebPtrs() function
		if (!(LpGetWebPtrs = (GetWebPtrsPtr *)GetProcAddress(CWebDll, GETWEBPTRSNAME)))
		{
			MessageBox(MainWindow, _T("Need a never version of CWebPage DLL!"), _T("ERROR"), MB_OK);
			FreeLibrary(CWebDll);
			CWebDll = 0;
		}

		LpWaitOnReadyState = (WaitOnReadyStatePtr *)GetProcAddress(CWebDll, WAITONREADYSTATENAME);
		LpGetWebElement = (GetWebElementPtr *)GetProcAddress(CWebDll, GETWEBELEMENTNAME);
		LpCreateWebEvtHandler = (CreateWebEvtHandlerPtr *)GetProcAddress(CWebDll, CREATEWEBEVTHANDLERNAME);
		LpSetWebReturnValue = (SetWebReturnValuePtr *)GetProcAddress(CWebDll, SETWEBRETURNVALUENAME);
	}
	else
		MessageBox(MainWindow, _T("Can't open CWebPage DLL!"), _T("ERROR"), MB_OK);

	return(CWebDll);
}





/****************************** WinMain() ***************************
 * C program entry point.
 *
 * This creates a window to host the web browser, and displays a web
 * page.
 */

int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hInstNULL, LPSTR lpszCmdLine, int nCmdShow)
{
	// Load CWebPage.DLL and get ptrs to functions we need to call
	if (open_webdll())
	{
		{
		WNDCLASSEX		wc;

		// Register the class of our Main window. 'windowProc' is our message handler
		// and 'ClassName' is the class name. You can choose any class name you want.
		ZeroMemory(&wc, sizeof(WNDCLASSEX));
		wc.cbSize = sizeof(WNDCLASSEX);
		wc.hInstance = hInstance;
		wc.lpfnWndProc = windowProc;
		wc.lpszClassName = &ClassName[0];
		wc.style = CS_CLASSDC|CS_HREDRAW|CS_VREDRAW|CS_PARENTDC|CS_BYTEALIGNCLIENT|CS_DBLCLKS;
		RegisterClassEx(&wc);

		// Register the class of our window to host the browser. 'browserWindowProc' is our message handler
		// and 'BrowserClassName' is the class name. You can choose any class name you want.
		wc.lpfnWndProc = browserWindowProc;
		wc.lpszClassName = &BrowserClassName[0];
		wc.style = CS_HREDRAW|CS_VREDRAW;
		RegisterClassEx(&wc);
		}

		// Create a Main window.
		if ((MainWindow = CreateWindowEx(0, &ClassName[0], _T("A web browser"), WS_OVERLAPPEDWINDOW,
						CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
						HWND_DESKTOP, NULL, hInstance, 0)))
		{
			MSG			msg;

			// Create a child window to host the browser object. NOTE: We embed the browser object
			// duing our WM_CREATE handling for this child window.
			if ((msg.hwnd = CreateWindowEx(0, &BrowserClassName[0], 0, WS_CHILD|WS_VISIBLE|WS_CLIPSIBLINGS,
								0, 0, 100, 100,
								MainWindow, (HMENU)1000, hInstance, 0)))
			{
				register IHTMLElement *		elem;
				IHTMLDocument2 *			htmlDoc2;
				IHTMLWindow2 *				htmlWin2;
				VARIANT						varDisp;

				// Display an HTML string with 2 elements on it, and a link
				LpDisplayHTMLStr(msg.hwnd,
					_T("<FONT id=testfont color=red>This is a test of events</font>")
					_T("<FORM id=testform action=\"http://www.google.com/search\">")
					_T("<Input type=text name=testinput>")
					_T("<Input type=submit>")
					_T("</FORM><A HREF=\"http://www.microsoft.com\">Goto another page</A>")
					);
	
				// Wait for the browser to achieve its READYSTATE_COMPLETE (before we
				// try to get the IHTMLDocument2).
				if (LpWaitOnReadyState(msg.hwnd, READYSTATE_COMPLETE, 3000, NULL) != WORS_SUCCESS ||

					// Get the IHTMLDocument2 object.
					LpGetWebPtrs(msg.hwnd, 0, &htmlDoc2))
				{
					goto bad;
				}
	
				VariantInit(&varDisp);
				varDisp.vt = VT_DISPATCH;

				// Get the IHTMLElement object for the FONT tag (whose id is "testfont")
				if ((elem = LpGetWebElement(msg.hwnd, htmlDoc2, _T("testfont"), 0)) &&

					// Create one of our _IDispatchEx objects associated with this FONT tag. Messages
					// about events we "capture" will be sent to our window (by _IDispatchEx's Invoke()).
					// Note that we'll store our IHTMLElement object into our _IDispatchEx's "object"
					// field. We arbitrarily assign an ID of 1 to this FONT tag's events. We'll use this ID
					// in our WM_NOTIFY handling to refer to events from this FONT tag. So, we pass that as
					// the ID argument
					(varDisp.pdispVal = LpCreateWebEvtHandler(msg.hwnd, htmlDoc2, 0, 1, (IUnknown *)elem, 0)))
				{
					// "Capture" the "mouseover" event. When the user moves the mouse over
					// this FONT tag, our _IDispatchEx's Invoke() will be called, and it will
					// in turn send a WM_NOTIFY to our window
					elem->lpVtbl->put_onmouseover(elem, varDisp);

					// "Capture" the "mouseout" event
					elem->lpVtbl->put_onmouseout(elem, varDisp);

					// NOTE: We must not VariantClear(&varDisp) or that will deallocate our _IDispatchEx, and
					// we don't want to do that now! Also, we don't Release "elem" because we told 
					// CreateWebEvtHandler to store that in the returned _IDispatchEx. We'll need it later
					// when we receive a WM_NOTIFY message with this _IDispatchEx.
				}

				// Now do a similiar thing for the FORM. But in this case we capture the
				// "submit" event. We also assign this element an ID of 2. We store
				// an IHTMLFormElement object in the _IDispatchEx's object field. NOTE:
				// the FORM has an id name of "testform"
				if ((elem = LpGetWebElement(msg.hwnd, htmlDoc2, _T("testform"), 0)))
				{
					IHTMLFormElement *	htmlForm;

					if (!elem->lpVtbl->QueryInterface(elem, &IID_IHTMLFormElement, (void **)&htmlForm))
					{
						if ((varDisp.pdispVal = LpCreateWebEvtHandler(msg.hwnd, htmlDoc2, 0, 2, (IUnknown *)htmlForm, 0)))
							htmlForm->lpVtbl->put_onsubmit(htmlForm, varDisp);
					}
					else
						elem->lpVtbl->Release(elem);
				}

				// Capture any "scroll" event. This event is not associated with any particular element
				// on a web page. So unlike the above, where we get the HTMLElement object for the element,
				// we don't do that here. Instead, we need to get the IHTMLWindow2 object and use that.
				// Also note that we arbitrary assign an ID of 3 to these events
				if (!htmlDoc2->lpVtbl->get_parentWindow(htmlDoc2, &htmlWin2) && htmlWin2 &&

					(varDisp.pdispVal = LpCreateWebEvtHandler(msg.hwnd, htmlDoc2, 0, 3, (IUnknown *)htmlWin2, 0)))
				{
					htmlWin2->lpVtbl->put_onscroll(htmlWin2, varDisp);
				}

				// Capture any "dblclick" event. Again, this is not associated with any particular element.
				// We arbitrarily assign an ID of 4 to these events
				if ((varDisp.pdispVal = LpCreateWebEvtHandler(msg.hwnd, htmlDoc2, 0, 4, (IUnknown *)htmlDoc2, 0)))
				{
					htmlDoc2->lpVtbl->put_ondblclick(htmlDoc2, varDisp);
				}

				// NOTE: We don't free the IHTMLDocument2 we got above because we told CreateWebEvtHandler
				// to store it for the double-click event above. If we had several events that needed
				// this pointer, then we'd call GetWebPtrs(hwnd, 0, &htmlDoc2) for each one.

				// Show the Main window.
				ShowWindow(MainWindow, nCmdShow);
				UpdateWindow(MainWindow);

				// Do a message loop until WM_QUIT.
				while (GetMessage(&msg, 0, 0, 0) == 1)
				{
					TranslateMessage(&msg);
					DispatchMessage(&msg);
				}
			}
			else
			{
				MessageBox(MainWindow, _T("Can't create browser child window!"), _T("ERROR"), MB_OK);
	bad:		DestroyWindow(MainWindow);
			}
		}

		// Free the DLL.
		FreeLibrary(CWebDll);
	}

	return(0);
}
