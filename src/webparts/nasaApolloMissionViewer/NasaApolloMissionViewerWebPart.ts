import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NasaApolloMissionViewerWebPart.module.scss';
import * as strings from 'NasaApolloMissionViewerWebPartStrings';

import { IMission } from "../../models";

import { MissionService } from "../../services";


export interface INasaApolloMissionViewerWebPartProps {
  description: string;
}

export default class NasaApolloMissionViewerWebPart extends BaseClientSideWebPart<INasaApolloMissionViewerWebPartProps> {

  private selectedMission: IMission = this._getSelectedMission();
  private missionDetailElement: HTMLElement;


  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.nasaApolloMissionViewer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
              <div class="apolloMissionDetails"></div>
            </div>
          </div>
        </div>
      </div>`;


      this.missionDetailElement = this.domElement.getElementsByClassName('apolloMissionDetails')[0] as HTMLElement;
      if(this.selectedMission)
      {
        this._renderMissionDetails(this.missionDetailElement, this.selectedMission);
      } else {
        this.missionDetailElement.innerHTML = '';
      }
  }

  private _getSelectedMission(): IMission {
    const selectedMissionId: string = 'AS-506';
    return MissionService.getMission(selectedMissionId);
  }

  private _renderMissionDetails(element: HTMLElement, mission: IMission): void {
    element.innerHTML = `
      <p class="ms-font-m">
        <span class="ms-fontWeight-semibold">Mission: </span>
        ${escape(mission.name)}
      </p>
      <p class="ms-font-m">
        <span class="ms-fontWeight-semibold">Duration: </span>
        ${escape(this._getMissionTimeLine(mission))}
      </p>
      <a href="${mission.wiki_href}" target="_blank" class="${styles.button}">
      <span class="${styles.label}">Learn more about ${escape(mission.name)} on Wikipedia &raquo;</span>
      </a>`;
  }

  private _getMissionTimeLine(mission: IMission): string {
    let missionDate = mission.end_date !== ''
    ? `${mission.launch_date.toString()} - ${mission.end_date.toString()}`
    : `${mission.launch_date.toString()}`;
    return missionDate;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
