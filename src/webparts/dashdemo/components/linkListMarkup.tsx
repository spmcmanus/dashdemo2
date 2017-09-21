// React
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Iframe from 'react-iframe';

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
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

// Custom components and properties
import { IDashdemoProps } from './IDashdemoProps';

// Styling
import styles from '../resources/Dashdemo.module.scss';

// local state
export interface localState {
    rowClasses: string;
}

// component class definition
export default class LinkListContainer extends React.Component<any, localState> {

    // constructor
    public constructor(props: IDashdemoProps, state: localState) {
        super(props);
    }

    public getImageSrc(author) {
        const filename = author.replace(/\s+/g, '') + ".jpg";
        const fullUrl = window.location.origin + "/sites/dev/SiteAssets/" + filename;
        return fullUrl;
    }
    // return loading if the incidents state has not yet been set
    public render(): React.ReactElement<IDashdemoProps> {
        const links = this.props.links.slice(0, this.props.showRecentIncidents);
        const handler = this.props.handler;
        const rootUrl = window.location.origin;
        const siteName = "dev";
        const listName = "SiteAssets";
        const fileName = "preview200.jpg";
        let previewURL = rootUrl + "/sites/" + siteName + "/" + listName + "/" + fileName;

        if (!links) {
            return <div>Loading...</div>;
        }
        // return list of incidents
        return (
            <div className="ms-Grid">
                <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm12">
                        {links.map((link, key) => {
                            if (link != null) {
                                const date = new Date(link.Created);
                                var formatOptions = {
                                    day: '2-digit',
                                    month: '2-digit',
                                    year: 'numeric',
                                    hour: '2-digit',
                                    minute: '2-digit',
                                    hour12: true
                                };
                                const displayDate = date.toLocaleDateString('en-US', formatOptions);

                                if (link.incidentPhotos != null) {
                                    previewURL = link.incidentPhotos.Url;
                                }
                                const thisPreviewProps: IDocumentCardPreviewProps = {
                                    previewImages: [
                                        {
                                            previewImageSrc: previewURL,
                                            imageFit: ImageFit.none
                                        }
                                    ],
                                };
                                var fileExt = null;
                                var fileExtClass = null;
                                var fileExtClassSm = null;
                                if (link.AttachmentFiles.results.length > 0) {
                                    fileExt = link.AttachmentFiles.results[0].FileName.split(".")[1].trim();
                                    if (fileExt == 'xlsx' || fileExt == 'xls') {
                                        fileExtClass = 'ms-Icon ms-Icon--ExcelLogo ms-fontSize-su ms-fontColor-green';
                                        fileExtClassSm = 'ms-Icon ms-Icon--ExcelLogo ms-fontSize-1 ms-fontColor-green';
                                    } else if (fileExt == 'docx' || fileExt == 'doc') {
                                        fileExtClass = 'ms-Icon ms-Icon--WordLogo ms-fontSize-su ms-fontColor-blue';
                                        fileExtClassSm = 'ms-Icon ms-Icon--WordLogo ms-fontSize-1 ms-fontColor-blue';
                                    } else if (fileExt == 'pptx' || fileExt == 'ppt') {
                                        fileExtClass = 'ms-Icon ms-Icon--PowerPointLogo ms-fontSize-su ms-fontColor-redDark';
                                        fileExtClassSm = 'ms-Icon ms-Icon--PowerPointLogo ms-fontSize-1 ms-fontColor-redDark';
                                    } else if (fileExt == 'pdf') {
                                        fileExtClass = 'ms-Icon ms-Icon--PDF ms-fontSize-su ms-fontColor-red';
                                        fileExtClassSm = 'ms-Icon ms-Icon--PDF ms-fontSize-l ms-fontColor-red';
                                    } else {
                                        fileExtClass = 'ms-Icon ms-Icon--Page ms-fontSize-su ms-fontColor-blue';
                                        fileExtClassSm = 'ms-Icon ms-Icon--Page ms-fontSize-l ms-fontColor-blue';
                                    }
                                } else if (link.linkURL != '') {
                                    fileExtClass = "ms-Icon ms-Icon--Website ms-fontSize-su ms-fontColor-blue";
                                    fileExtClassSm = "ms-Icon ms-Icon--Website ms-fontSize-l ms-fontColor-blue";
                                } else {
                                    fileExtClass = "";
                                    fileExtClassSm = "";
                                }
                                if (this.props.cardType == 0) {
                                    return (
                                        <DocumentCard
                                            className={styles.documentCard}
                                            onClick={handler.bind(this, link)}>
                                            <div className={styles.iconContainer}>
                                                <div className={fileExtClass}></div>
                                            </div>
                                            <DocumentCardTitle
                                                title={link.Title}
                                                shouldTruncate={true}
                                            />
                                            <DocumentCardActivity
                                                activity={displayDate}
                                                people={[{
                                                    name: link.documentAuthor,
                                                    profileImageSrc: this.getImageSrc(link.documentAuthor)
                                                }]}
                                            />
                                            <DocumentCardActions
                                                actions={
                                                    [
                                                        {
                                                            iconProps: { iconName: 'Share' },
                                                            onClick: (ev: any) => {
                                                                ev.preventDefault();
                                                                ev.stopPropagation();
                                                            },
                                                            ariaLabel: 'share action'
                                                        },
                                                        {
                                                            iconProps: { iconName: 'Pin' },
                                                            onClick: (ev: any) => {
                                                                ev.preventDefault();
                                                                ev.stopPropagation();
                                                            },
                                                            ariaLabel: 'pin action'
                                                        },
                                                        {
                                                            iconProps: { iconName: 'Ringer' },
                                                            onClick: (ev: any) => {
                                                                ev.preventDefault();
                                                                ev.stopPropagation();
                                                            },
                                                            ariaLabel: 'ringer action'
                                                        },
                                                    ]
                                                }
                                                views={432}
                                            />
                                        </DocumentCard>
                                    );
                                } else {
                                    let previewProps: IDocumentCardPreviewProps = {
                                        previewImages: [{
                                            name: 'Revenue stream proposal fiscal year 2016 version02.pptx',
                                            url: 'http://bing.com',
                                            iconSrc: 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-react-assets/icon-ppt.png',
                                            width: 24
                                        }]
                                    };
                                    return (
                                        <DocumentCard
                                            type={DocumentCardType.compact}
                                            className={styles.documentCard}
                                            onClick={handler.bind(this, link)}>
                                            <div className='ms-DocumentCard-details'>
                                            <div className={styles.iconContainer}>
                                            <div className={[fileExtClassSm, styles.inline].join(' ')}></div>
                                                <div className={[styles.smallTitle,styles.inline].join(' ')}>{link.Title}</div>
                                            </div>
                                            <DocumentCardActivity
                                                activity={displayDate}
                                                people={[{
                                                    name: link.documentAuthor,
                                                    profileImageSrc: this.getImageSrc(link.documentAuthor)
                                                }]}
                                            />
                                            </div>
                                        </DocumentCard>
                                    );
                                }
                            }
                        })}
                    </div>
                </div>
            </div>
        );
    }
}