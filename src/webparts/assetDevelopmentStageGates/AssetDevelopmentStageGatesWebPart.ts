import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './AssetDevelopmentStageGatesWebPart.module.scss';
import * as strings from 'AssetDevelopmentStageGatesWebPartStrings';

export interface IAssetDevelopmentStageGatesWebPartProps {
  description: string;
}

export default class AssetDevelopmentStageGatesWebPart extends BaseClientSideWebPart<IAssetDevelopmentStageGatesWebPartProps> {

  public render(): void {
    try {
      this.domElement.innerHTML = `
        <section class="${styles.assetDevelopmentStageGates}">
          <div class="${styles.page}">
            <div class="${styles.topTabs}">
              <div class="${styles.tabGroup}">
                <div class="${styles.tabRow}">
                  <div class="${styles.tab}">Permitting Office scope &amp; mandate</div>
                </div>
                <div class="${styles.tabRow} ${styles.tabDivider}">
                  <div class="${styles.tab}">Land Planning</div>
                </div>
              </div>

              <div class="${styles.tabGroup}">
                <div class="${styles.tabRow}">
                  <div class="${styles.tab}">Permitting process execution</div>
                </div>
                <div class="${styles.tabRow} ${styles.tabDivider}">
                  <div class="${styles.tab}">Land Servicing</div>
                </div>
              </div>

              <div class="${styles.tabGroup}">
                <div class="${styles.tabRow}">
                  <div class="${styles.tab} ${styles.activeBlue}">Mapping to Stage Gates</div>
                </div>
                <div class="${styles.tabRow} ${styles.tabDivider}">
                  <div class="${styles.tab} ${styles.activeGold}">Asset Development</div>
                </div>
              </div>
            </div>

            <div class="${styles.titleSection}">
              <h1>
                <strong>Asset Development Stage Gates:</strong> <span>Key external</span><br>
                <span>approvals, permits and NOCs mapped to the process</span>
                <span class="${styles.nonExhaustive}">Non-exhaustive</span>
              </h1>
            </div>

            <div class="${styles.stageMappingWrap}">
              <div class="${styles.stageHeaderRow}">
                <div class="${styles.stageHeaderCell}">Stage 1</div>
                <div class="${styles.stageHeaderCell}">Stage 2</div>
                <div class="${styles.stageHeaderCell}">Stage 3</div>
                <div class="${styles.stageHeaderCell}">Stage 4</div>
                <div class="${styles.stageHeaderCell}">Stage 5</div>
                <div class="${styles.stageHeaderCell}">Stage 6</div>
                <div class="${styles.stageHeaderCell} ${styles.split}">
                  <div class="${styles.splitTop}">Stage 7</div>
                  <div class="${styles.splitBottom}">
                    <span>Stage 7.1</span>
                    <span>Stage 7.2</span>
                  </div>
                </div>
              </div>

              <div class="${styles.stageNameRow}">
                <div class="${styles.stageNameCell}">Strategy Brief</div>
                <div class="${styles.stageNameCell}">Development Brief &amp; Pre-Concept Design</div>
                <div class="${styles.stageNameCell}">Concept Design</div>
                <div class="${styles.stageNameCell}">Schematic Design</div>
                <div class="${styles.stageNameCell}">Detailed Design</div>
                <div class="${styles.stageNameCell}">Construction Commencement</div>
                <div class="${styles.stageNameCell} ${styles.splitName}">
                  <span>Practical Completion</span>
                  <span>Pre-Operations</span>
                </div>
              </div>

              <div class="${styles.contentGrid}">
                <div class="${styles.col}">
                  <div class="${styles.stage1Text}">
                    At this stage<br>Project Team<br>needs to<br>
                    <strong><em>identify all<br>external NOCs<br>and permits</em></strong>
                    needed for the<br>project with its<br>AOR
                  </div>
                </div>

                <div class="${styles.col}"></div>

                <div class="${styles.col}">
                  ${this.renderCard('1', 'TIS: Scope & Preliminary Approval', styles.cardApproval)}
                  ${this.renderCard('2<br>–<br>4', 'PoC request for Electricity, Water and Wastewater Approval', `${styles.cardApproval} ${styles.cardApprovalRange}`)}
                  ${this.renderCard('5', 'Telecom Infrastructure Compliance & Service Tie-In Approval', styles.cardApproval)}
                  ${this.renderCard('6', 'Heritage & Archaeological Survey Validation and Design Integration Approval', styles.cardApproval)}
                </div>

                <div class="${styles.col}">
                  ${this.renderCard('1', 'Fire & Life Safety Strategy Review (Concept) Approval', styles.cardApproval)}
                  ${this.renderCard('2', 'Preliminary Stormwater Outfall & Drainage Tie-In Approval', styles.cardApproval)}
                  ${this.renderCard('3', 'Building Permit', styles.cardPermit)}
                </div>

                <div class="${styles.col}">
                  ${this.renderCard('1', 'Civil Defense (Fire & Life Safety) Approval', styles.cardApproval)}
                  ${this.renderCard('2', 'Final Utility Connection NOC', styles.cardNoc)}
                  ${this.renderCard('3', 'Environmental Permit to Construct', styles.cardPermit)}
                  ${this.renderCard('4', 'Obstacle Limitation Surface Approval', styles.cardApproval)}
                  ${this.renderCard('5', 'Height Clearance Approval', styles.cardApproval)}
                  ${this.renderCard('6', 'MOI Approval Letter on CCTV Designs', styles.cardApproval)}
                </div>

                <div class="${styles.col}">
                  ${this.renderCard('1', 'Site Preparation Approval', styles.cardApproval)}
                  ${this.renderCard('2', 'Excavation License Approval', styles.cardApproval)}
                  ${this.renderCard('3', 'Road Works Permit', styles.cardPermit)}
                  ${this.renderCard('4', 'Traffic Diversion Permit', styles.cardPermit)}
                </div>

                <div class="${styles.col} ${styles.splitCol}">
                  <div class="${styles.subCol}">
                    ${this.renderCard('1', 'Utility Service Activation Notification Approval', styles.cardApproval)}
                    ${this.renderCard('2', 'Civil Defense Occupancy Certificate Approval', styles.cardApproval)}
                  </div>

                  <div class="${styles.subCol}">
                    ${this.renderCard('1', 'Environmental Permit to Operate', styles.cardPermit)}
                    ${this.renderCard('2', 'Occupancy Certificate Approval', styles.cardApproval)}
                    ${this.renderCard('3', 'MOI Completion Certificate on CCTV Designs (As-Built) Approval', styles.cardApproval)}
                  </div>
                </div>
              </div>
            </div>

            <div class="${styles.aorRow}">
              <div class="${styles.aorTriangle}"></div>
              <div class="${styles.aorText}">Engage an AoR before the project starts</div>
            </div>

            <div class="${styles.footerNote}">
              Note: At each stage, the project team must prepare and submit an application for external authorization and obtain approval before advancing to the next stage
            </div>

            <div class="${styles.legendRow}">
              <div class="${styles.legendItem}">
                <div class="${styles.legendBox} ${styles.permit}">Permit</div>
              </div>
              <div class="${styles.legendItem}">
                <div class="${styles.legendBox} ${styles.approval}">Approval</div>
              </div>
              <div class="${styles.legendItem}">
                <div class="${styles.legendBox} ${styles.noc}">NOC</div>
              </div>
            </div>
          </div>
        </section>`;
    } catch (error) {
      console.error('AssetDevelopmentStageGates render error:', error);

      let errorMessage = 'Unknown render error';

      if (error instanceof Error) {
        errorMessage = error.message;
      } else if (typeof error === 'string') {
        errorMessage = error;
      } else {
        try {
          errorMessage = JSON.stringify(error);
        } catch {
          errorMessage = String(error);
        }
      }

      this.domElement.innerHTML = `
        <div class="${styles.errorState}">
          Unable to render Asset Development Stage Gates web part.<br>
          Error: ${errorMessage}
        </div>`;
    }
  }

  private renderCard(num: string, text: string, extraClass: string): string {
    return `
      <a href="#" target="_blank" rel="noopener noreferrer" class="${styles.card} ${extraClass}">
        <span class="${styles.cardNum}">${num}</span>
        <span class="${styles.cardText}">${text}</span>
      </a>
    `;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || '');
      this.domElement.style.setProperty('--link', semanticColors.link || '');
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || '');
    }
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