// React
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Iframe from 'react-iframe';
// Styling
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
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
// Custom components and properties
import { IDashdemoProps } from './IDashdemoProps';

// local state
export interface localState {
	rowClasses: string;
}

// component class definition
export default class LinkListContainer extends React.Component<any, localState> {

	// constructor
	public constructor(props: IDashdemoProps, state: localState) {
		super(props);
		//	this.state = {
		//		rowClasses: "ms-Grid-col ms-sm12"
		//	};
	}

	public getImageSrc(author) {
		const filename = author.replace(/\s+/g, '') + ".jpg";
		const fullUrl = window.location.origin + "/sites/dev/SiteAssets/" + filename;
		return fullUrl
	}

	// return loading if the incidents state has not yet been set
	public render(): React.ReactElement<IDashdemoProps> {
		console.log('linkCard - render - this.props', this.props)
		const links = this.props.links.slice(0, this.props.showRecentIncidents);
		const handler = this.props.handler;
		const rootUrl = window.location.origin;
		const siteName = "dev"
		const listName = "SiteAssets"
		const fileName = "preview200.jpg"
		let previewURL = rootUrl + "/sites/" + siteName + "/" + listName + "/" + fileName;

		if (!links) {
			return <div>Loading...</div>;
		}
		// return list of incidents
		return (
			//<div className={styles.panelStyle} >
			//	<div className={styles.tableStyle} >
			<div className="ms-Grid">
				<div className="ms-Grid-row">
					<div className="ms-Grid-col ms-sm12">
						{links.map((link, key) => {

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
							if (this.props.cardType == 0) {
								return (
									<DocumentCard
										className={styles.documentCard}
										onClick={handler.bind(this, link)}>
										<DocumentCardPreview { ...thisPreviewProps } />
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
															console.log('You clicked the share action.');
															ev.preventDefault();
															ev.stopPropagation();
														},
														ariaLabel: 'share action'
													},
													{
														iconProps: { iconName: 'Pin' },
														onClick: (ev: any) => {
															console.log('You clicked the pin action.');
															ev.preventDefault();
															ev.stopPropagation();
														},
														ariaLabel: 'pin action'
													},
													{
														iconProps: { iconName: 'Ringer' },
														onClick: (ev: any) => {
															console.log('You clicked the ringer action.');
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
								return (
									<DocumentCard
										type={DocumentCardType.compact}
										className={styles.documentCard}
										onClick={handler.bind(this, link)}>
										<div className='ms-DocumentCard-details'>
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
										</div>
									</DocumentCard>
								);
							}
						})}
					</div>
				</div>
			</div>
			//	</div>
			//	</div >
		);
	}
}