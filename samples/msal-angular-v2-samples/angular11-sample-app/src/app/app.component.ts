import { HttpClient } from '@angular/common/http';
import { Component, OnInit, Inject, OnDestroy } from '@angular/core';
import { MsalService, MsalBroadcastService, MSAL_GUARD_CONFIG, MsalGuardConfiguration } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType, InteractionType, PopupRequest, RedirectRequest } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

declare const OfficeRuntime: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit, OnDestroy {
  title = 'Msal2 sample';
  isIframe = false;
  loggedIn = false;
  private readonly _destroying$ = new Subject<void>();

  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService,
    private http: HttpClient
  ) { }

  ngOnInit(): void {
    this.isIframe = window !== window.parent && !window.opener; // Remove this line to use Angular Universal

    this.checkAccount() || this.login();

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS || msg.eventType === EventType.ACQUIRE_TOKEN_SUCCESS),
        takeUntil(this._destroying$)
      )
      .subscribe((result) => {
        this.checkAccount();
      });
  }

  checkAccount() {
    console.log(JSON.stringify(this.authService.instance.getAllAccounts()));
    return this.loggedIn = this.authService.instance.getAllAccounts().length > 0;
  }

  login() {
    if (this.msalGuardConfig.interactionType === InteractionType.Popup) {
      if (this.msalGuardConfig.authRequest) {
        this.authService.loginPopup({ ...this.msalGuardConfig.authRequest } as PopupRequest)
          .subscribe((response: AuthenticationResult) => {
            this.authService.instance.setActiveAccount(response.account);
            this.checkAccount();
          });
      } else {
        this.authService.loginPopup()
          .subscribe((response: AuthenticationResult) => {
            this.authService.instance.setActiveAccount(response.account);
            this.checkAccount();
          });
      }
    } else {
      if (this.msalGuardConfig.authRequest) {
        this.authService.loginRedirect({ ...this.msalGuardConfig.authRequest } as RedirectRequest);
      } else {
        // if (OfficeRuntime) {
        //   OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false, allowConsentPrompt: false }).then((bootstrap: any) => {
        //     const body = { bstoken: bootstrap, scopes: ["Files.ReadWrite.All", "Group.Read.All", "Sites.Read.All", "OpenId", "Profile", "User.Read"].join(' ') };
        //     this.http.post('http://localhost:4199/api/token' , body).subscribe(resp => {
        //       console.log(resp);
        //     });
        //   })
        // } else {
          this.authService.loginRedirect({
            scopes: ["Files.ReadWrite.All", "Group.Read.All", "Sites.Read.All", "OpenId", "Profile", "User.Read"]
          });
        // }
      }
    }
  }

  logout() {
    this.authService.logout();
  }

  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
  }
}
