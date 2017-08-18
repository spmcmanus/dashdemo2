import * as React from 'react';
import * as jquery from 'jquery';

import { IDashdemoProps } from './IDashdemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import Iframe from 'react-iframe';

// styling
import styles from '../resources/Dashdemo.module.scss';

// Office-Ui Fabric Components
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardActions,
  IDocumentCardPreviewProps,
  DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';

import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

import LinkListMarkup from './linkListMarkup';

export interface linksState {
  links: [
    {
      "Title": string;
      "AuthorId": string;
      "linkURL": string;
      "linkDesc": string;
    }
  ];
  linkSelectedURL: string;
  rowClasses: string;
  embedClasses: string;
  cardType: number;
}

export default class Dashdemo extends React.Component<IDashdemoProps, linksState> {

  public constructor(props: IDashdemoProps, state: linksState) {
    super(props);

    console.log('document card dtype', DocumentCardType.compact)


    this.state = {
      links:
      [{
        "Title": '',
        "AuthorId": '',
        "linkURL": '',
        "linkDesc": '',
      }],
      linkSelectedURL: "",
      rowClasses: "ms-Grid-col ms-sm12",
      embedClasses: "",
      cardType: DocumentCardType.normal
    };
    this.onCardClick = this.onCardClick.bind(this);
  }

  // card click listener
  public onCardClick(link, e) {
    console.log("clicky")
    console.log("thisLink", link)
    if (link.AttachmentFiles.results[0] == undefined) {
      console.log("display link instead");
      var win = window.open(link.linkURL, '_blank');
      win.focus();
      //
    } else {
      console.log("work off attachment")
      const linkId = link.ID;
      const fileName = link.AttachmentFiles.results[0].FileName;
      const fileExt = fileName.substr(fileName.lastIndexOf('.') + 1);
      if (fileExt == 'docx' || fileExt == 'doc' || fileExt == 'xlsx' || fileExt == 'pptx') {
        console.log("office doc...embed")
        const attachmentURLRoot = window.location.origin + "/sites/dev/_layouts/15/WopiFrame.aspx?sourcedoc=";
        const attachmentURL = attachmentURLRoot + "/sites/dev/Lists/DashboardLinks/Attachments/" + linkId + "/" + fileName;
        const extras = "&action=embedview&wdbipreview=true";
        const attachmentFullURL = attachmentURL;
        console.log("attachmentFullURL", attachmentFullURL);
        this.setState({
          links: this.state.links,
          linkSelectedURL: attachmentFullURL,
          rowClasses: 'ms-Grid-col ms-sm4',
          embedClasses: "",
          cardType: DocumentCardType.compact
        });
      } else {
        console.log("non office document ... download");
        var win = window.open(link.AttachmentFiles.results[0].ServerRelativeUrl, '_blank');
        win.focus();
      }
    }
    console.log("click is done")
  }

  public toggleCardDisplay() {
    console.log('toggling card type')
    if (this.state.cardType == DocumentCardType.normal) {
      this.setState({
        links: this.state.links,
        linkSelectedURL: this.state.linkSelectedURL,
        rowClasses: this.state.rowClasses,
        embedClasses: this.state.embedClasses,
        cardType: DocumentCardType.compact
      });
    } else {
      this.setState({
        links: this.state.links,
        linkSelectedURL: this.state.linkSelectedURL,
        rowClasses: this.state.rowClasses,
        embedClasses: this.state.embedClasses,
        cardType: DocumentCardType.normal
      });
    }
  }

  public clearSelected() {
    this.setState({
      links: this.state.links,
      linkSelectedURL: '',
      rowClasses: this.state.rowClasses,
      embedClasses: this.state.embedClasses,
      cardType: DocumentCardType.normal
    });
  }

  public componentDidMount() {
    var reactHandler = this;
    const rootUrl = window.location.origin;
    const listName = "DashboardLinks";
    const siteName = "dev"
    // const fullUrl = rootUrl + "/sites/" + siteName + "/_api/web/lists/GetByTitle('" + listName + "')/Items"//(1)/AttachmentFiles";
    const fullUrl = rootUrl + "/sites/" + siteName + "/_api/web/lists/GetByTitle('" + listName + "')/Items?$expand=AttachmentFiles";


    jquery.ajax({
      url: fullUrl,
      type: "GET",
      dataType: "json",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        console.log(resultData)
        reactHandler.setState({
          links: resultData.d.results,
          linkSelectedURL: '',
          rowClasses: 'ms-Grid-col ms-sm12',
          embedClasses: '',
          cardType: DocumentCardType.normal
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log('jqXHR', jqXHR);
        console.log('text status', textStatus);
        console.log('error', errorThrown);
      }
    });
  }

  public render(): React.ReactElement<IDashdemoProps> {
    console.log("state", this.state)
    if (this.state.links[0].Title == '') {
      return (
        <div>Loading...</div>
      );
    } else if (this.state.linkSelectedURL != '') {
      return (
        <div id="mainContainer">
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12">
                <div className={styles.buttonRight}>
                  <PrimaryButton
                    text='Back'
                    onClick={() => this.clearSelected()}
                  />
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm3">
                <LinkListMarkup
                  links={this.state.links}
                  rowClasses={this.state.rowClasses}
                  embed={this.state.embedClasses}
                  selected={this.state.linkSelectedURL}
                  handler={this.onCardClick}
                  cardType={this.state.cardType}
                ></LinkListMarkup>
              </div>
              <div className="ms-Grid-col ms-sm9">
                <div>
                  <Iframe url={this.state.linkSelectedURL}
                    width="100%"
                    height="1000px"
                    display="initial"
                    position="relative"
                    allowFullScreen>
                  </Iframe>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    } else {
      const theseIncidents = this.state.links;
      return (
        <div>
          <div className={styles.buttonRight}>
            <PrimaryButton
              text='Toggle Card Display'
              onClick={() => this.toggleCardDisplay()}
            />
          </div>
          <LinkListMarkup
            links={this.state.links}
            rowClasses={this.state.rowClasses}
            embed={this.state.embedClasses}
            selected={this.state.linkSelectedURL}
            handler={this.onCardClick}
            cardType={this.state.cardType}
          ></LinkListMarkup>
        </div>
      );
    }
  }
}
