package Win32::PowerPoint;

use strict;
use warnings;

our $VERSION = '0.01';

use File::Spec;
use Win32::OLE;
use Win32::PowerPoint::Constants;

sub new {
  my $class = shift;
  my $self  = bless {
    c            => Win32::PowerPoint::Constants->new,
    was_invoked  => 0,
    app          => undef,
    presentation => undef,
    slide        => undef,
  }, $class;

  $self->{app} = Win32::OLE->GetActiveObject('PowerPoint.Application');

  unless (defined $self->{app}) {
    $self->{app} = Win32::OLE->new('PowerPoint.Application')
      or die Win32::OLE->LastError;
    $self->{was_invoked} = 1;
  }

  return $self;
}

sub quit {
  my $self = shift;

  return unless $self->{app};

  $self->{app}->Quit;
  $self->{app} = undef;
}

sub new_presentation {
  my $self = shift;

  return unless $self->{app};

  $self->{slide} = undef;

  $self->{presentation} = $self->{app}->Presentations->Add
    or die Win32::OLE->LastError;
}

sub save_presentation {
  my $self = shift;
  my $file = shift;

  return unless $self->{presentation};

  $self->{presentation}->SaveAs( File::Spec->rel2abs($file) );
}

sub new_slide {
  my $self = shift;

  $self->{slide} = $self->{presentation}->Slides->Add(
    $self->{presentation}->Slides->Count + 1,
    $self->{c}->LayoutBlank
  ) or die Win32::OLE->LastError;
}

sub add_text {
  my ($self, $text, $option) = @_;

  return unless $self->{slide};

  $option = {} unless ref $option eq 'HASH';

  $text =~ s/\n/\r/gs;

  my $num_of_boxes = $self->{slide}->Shapes->Count;
  my $last   = $self->{slide}->Shapes($num_of_boxes);
  my ($left, $top, $width, $height);
  if ($last) {
    $left   = $option->{left}   || $last->Left;
    $top    = $option->{top}    || $last->Top + $last->Height + 20;
    $width  = $option->{width}  || $last->Width;
    $height = $option->{height} || $last->Height;
  }
  else {
    $left   = $option->{left}   || 30;
    $top    = $option->{top}    || 30;
    $width  = $option->{width}  || 600;
    $height = $option->{height} || 200;
  }

  my $new_box = $self->{slide}->Shapes->AddTextbox(
    $self->{c}->TextOrientationHorizontal,
    $left, $top, $width, $height
  );

  my $frame = $new_box->TextFrame;
  my $range = $frame->TextRange;

  $frame->{WordWrap} = $self->{c}->True;
  $range->ParagraphFormat->{Alignment} = $self->{c}->AlignLeft;
  $range->ParagraphFormat->{FarEastLineBreakControl} = $self->{c}->True;
  $range->{Text} = $text;

  $range->Font->{Bold}      = $self->{c}->True if $option->{bold};
  $range->Font->{Italic}    = $self->{c}->True if $option->{italic};
  $range->Font->{Underline} = $self->{c}->True if $option->{underline};
  $range->Font->{Size}      = $option->{size}  if $option->{size};

  $frame->{AutoSize} = $self->{c}->AutoSizeNone;
  $frame->{AutoSize} = $self->{c}->AutoSizeShapeToFitText;
}

sub close_presentation {
  my $self = shift;

  $self->{presentation}->Close;
  $self->{presentation} = undef;
}

sub DESTROY {
  my $self = shift;

  $self->quit if $self->{was_invoked};
}

1;
__END__

=head1 NAME

Win32::PowerPoint - helps to convert texts to PP slides

=head1 SYNOPSIS

    use Win32::PowerPoint;

    # invoke (or connect to) PowerPoint
    my $pp = Win32::PowerPoint->new;

    $pp->new_presentation;

    ... (load and parse your slide text)

    foreach my $slide (@slides) {
      $pp->new_slide;

      $pp->add_text($slide->title, { size => 40, bold => 1 });
      $pp->add_text($slide->body);
    }

    $pp->save_presentation('slide.ppt');

    $pp->close_presentation;

    # PowerPoint closes automatically

=head1 DESCRIPTION

Win32::PowerPoint mainly aims to help to convert L<Spork> (or Sporx)
texts to PowerPoint slides. Though there's no converter at the moment,
you can add texts to your new slides/presentation and save it. 

=head1 METHODS

=head2 new

Invokes (or connects to) PowerPoint.

=head2 quit

Explicitly disconnects and close PowerPoint this module (or you!) invoked.

=head2 new_presentation

Creates a new (probably blank) presentation.

=head2 save_presentation (path)

Saves the presentation to where you specified. Accepts relative path.
You might want to save it as .pps (slideshow) file to make it easy to
show slides (it just starts full screen slideshow with doubleclick).

=head2 close_presentation

Explicitly closes the presentation

=head2 new_slide

Adds a new (blank) slide to the presentation.

=head2 add_text (text, option)

Adds (formatted) text to the slide. Options are:

=over 4

=item left, top, width, height

of the Textbox.

=item bold, italic, underline, size

of the Text.

=back

=head1 TODO

A lot, but if you want to make the most of PowerPoint, use Win32::OLE
directly, or better, use PowerPoint itself (and maybe its VBA) ;)

=head1 AUTHOR

Kenichi Ishigaki, E<lt>ishigaki@cpan.orgE<gt>

=head1 COPYRIGHT AND LICENSE

Copyright (C) 2006 by Kenichi Ishigaki

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut

