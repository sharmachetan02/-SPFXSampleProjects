import * as React from 'react';
import { Icon } from './Interfaces';
import XIcon from './icons/X';
import FacebookIcon from './icons/Facebook';
import LinkedInIcon from './icons/LinkedIn';
import InstagramIcon from './icons/Instagram';
import YoutubeIcon from './icons/Youtube';
import PreviousArrowIcon from './icons/PreviousArrow';
import NextArrowIcon from './icons/NextArrow';
import HomeIcon from './icons/Home';
import NewsIcon from './icons/News';
import GroupIcon from './icons/Group';
import BlogIcon from './icons/Blog';
import OpenIcon from './icons/OpenIcon';
import CalendarIcon from './icons/Calendar';
import LocationIcon from './icons/Location';
import WatchIcon from './icons/Watch';
import StockExchangeIcon from './icons/StockExchange';
import TimeZoneIcon from './icons/TimeZone';
import DirectoryIcon from './icons/Directory';
import FlickrIcon from './icons/Flickr';
import ChevronDownIcon from './icons/ChevronDown';
import ChevronUpIcon from './icons/ChevronUp';
import CardIcon from './icons/Card';
import EmailIcon from './icons/Email';
import MobileIcon from './icons/Mobile';
import PhoneIcon from './icons/Phone';
import LocationCityIcon from './icons/LocationCity';
import DepartmentIcon from './icons/Department';
import RessourcesIcon from './icons/Ressources';
import ApplicationIcon from './icons/Application';
import RefreshIcon from './icons/Refresh';
import DPIcon from './icons/DP';
import EntityIcon from './icons/Entity';
import LegalSvcProviderIcon from './icons/LegalSvcProvider';
import SNPIcon from './icons/Snp';
import CountryIcon from './icons/Country';
import ContactIcon from './icons/Contact';
import TPIcon from './icons/TP';
import AuthorityIcon from './icons/Authority';
import MARequirementIcon from './icons/MARequirement';
import AgreementIcon from './icons/Agreement';
import BinocularsIcon from './icons/Binoculars';
import DishIcon from './icons/Dish';
import PlaneIcon from './icons/Plane';
import RocketIcon from './icons/Rocket';
import ShipIcon from './icons/Ship';
import TruckIcon from './icons/Truck';
import WifiIcon from './icons/Wifi';
import SatelliteIcon from './icons/Satellite';
import MarketIcon from './icons/Market';
import PaymentIcon from './icons/Payment';
import FeeIcon from './icons/Fee';

interface ISvgIconProps extends React.SVGProps<SVGSVGElement> {
  icon: Icon;
}

export default class SvgIcon extends React.Component<ISvgIconProps> {

  constructor(props) {
    super(props);
  }

  public async componentDidMount(): Promise<void> {
    //todo
  }

  public render(): React.ReactElement<ISvgIconProps> {
    const { icon } = this.props;
    switch (icon) {
      case Icon.X: {
        return <XIcon {...this.props} />;
      }
      case Icon.Facebook: {
        return <FacebookIcon {...this.props} />;
      }
      case Icon.Youtube: {
        return <YoutubeIcon {...this.props} />;
      }
      case Icon.LinkedIn: {
        return <LinkedInIcon {...this.props} />;
      }
      case Icon.Instagram: {
        return <InstagramIcon {...this.props} />;
      }
      case Icon.Flickr: {
        return <FlickrIcon {...this.props} />;
      }
      case Icon.PreviousArrow: {
        return <PreviousArrowIcon {...this.props} />;
      }
      case Icon.NextArrow: {
        return <NextArrowIcon {...this.props} />;
      }
      case Icon.Home: {
        return <HomeIcon {...this.props} />;
      }
      case Icon.News: {
        return <NewsIcon {...this.props} />;
      }
      case Icon.Group: {
        return <GroupIcon {...this.props} />;
      }
      case Icon.Blog: {
        return <BlogIcon {...this.props} />;
      }
      case Icon.Open: {
        return <OpenIcon {...this.props} />;
      }
      case Icon.Calendar: {
        return <CalendarIcon {...this.props} />;
      }
      case Icon.Location: {
        return <LocationIcon {...this.props} />;
      }
      case Icon.Watch: {
        return <WatchIcon {...this.props} />;
      }
      case Icon.StockExchange: {
        return <StockExchangeIcon {...this.props} />;
      }
      case Icon.TimeZone: {
        return <TimeZoneIcon {...this.props} />;
      }
      case Icon.Directory: {
        return <DirectoryIcon {...this.props} />;
      }
      case Icon.ChevronUp: {
        return <ChevronUpIcon {...this.props} />;
      }
      case Icon.ChevronDown: {
        return <ChevronDownIcon {...this.props} />;
      }
      case Icon.Card: {
        return <CardIcon {...this.props} />;
      }
      case Icon.Email: {
        return <EmailIcon {...this.props} />;
      }
      case Icon.Mobile: {
        return <MobileIcon {...this.props} />;
      }
      case Icon.Phone: {
        return <PhoneIcon {...this.props} />;
      }
      case Icon.LocationCity: {
        return <LocationCityIcon {...this.props} />;
      }
      case Icon.Application: {
        return <ApplicationIcon {...this.props} />;
      }
      case Icon.Department: {
        return <DepartmentIcon {...this.props} />;
      }
      case Icon.Ressources: {
        return <RessourcesIcon {...this.props} />;
      }
      case Icon.Refresh: {
        return <RefreshIcon {...this.props} />;
      }
      case Icon.DistributionPartner: {
        return <DPIcon {...this.props} />;
      }
      case Icon.EutelsatEntity: {
        return <EntityIcon {...this.props} />;
      }
      case Icon.LegalSvcProvider: {
        return <LegalSvcProviderIcon {...this.props} />;
      }
      case Icon.SNP: {
        return <SNPIcon {...this.props} />;
      }
      case Icon.Country: {
        return <CountryIcon {...this.props} />;
      }
      case Icon.Contact: {
        return <ContactIcon {...this.props} />;
      }
      case Icon.TeleportPartner: {
        return <TPIcon {...this.props} />;
      }
      case Icon.Authority: {
        return <AuthorityIcon {...this.props} />;
      }
      case Icon.MARequirement: {
        return <MARequirementIcon {...this.props} />;
      }
      case Icon.Agreement: {
        return <AgreementIcon {...this.props} />;
      }
      case Icon.Binoculars: {
        return <BinocularsIcon {...this.props} />;
      }
      case Icon.Dish: {
        return <DishIcon {...this.props} />;
      }
      case Icon.Plane: {
        return <PlaneIcon {...this.props} />;
      }
      case Icon.Rocket: {
        return <RocketIcon {...this.props} />;
      }
      case Icon.Satellite: {
        return <SatelliteIcon {...this.props} />;
      }
      case Icon.Ship: {
        return <ShipIcon {...this.props} />;
      }
      case Icon.Truck: {
        return <TruckIcon {...this.props} />;
      }
      case Icon.Wifi: {
        return <WifiIcon {...this.props} />;
      }
      case Icon.Market: {
        return <MarketIcon {...this.props} />;
      }
      case Icon.Fee: {
        return <FeeIcon {...this.props} />;
      }
      case Icon.Payment: {
        return <PaymentIcon {...this.props} />;
      }
      default: return null;
    }
  }
}
