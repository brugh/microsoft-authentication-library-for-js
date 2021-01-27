import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';

declare const Word: any;
const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/sites/root';

type SiteType = {
  webUrl?: string,
  displayName?: string,
  siteCollection?: { hostname?: string }
}

@Component({
  selector: 'app-orders',
  templateUrl: './lazy.component.html'
})
export class LazyComponent implements OnInit {
  public root!: SiteType;

  constructor(private http: HttpClient) {}

  ngOnInit() {
    this.getRoot();
  }

  getRoot() {
    this.http.get(GRAPH_ENDPOINT)
      .subscribe(root => {
        this.root = root;
      });
  }

  runWord() {
    Word.run(async (context: any) => {
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      paragraph.font.color = "blue";
      await context.sync();
    });
  }
}
